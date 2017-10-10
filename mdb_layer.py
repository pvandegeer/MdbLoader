import pyodbc, datetime
from PyQt4.QtGui import QProgressBar
from PyQt4.QtCore import QVariant, Qt
from qgis.utils import iface, QgsMessageBar
from qgis.core import (QgsVectorLayer, QgsFeature, QgsGeometry, QgsPoint, QgsField,
                       QgsMapLayerRegistry, QgsFeatureRequest, QgsMessageLog)


logger = lambda msg: QgsMessageLog.logMessage(msg, 'Mdb Layer', 1)

SHOW_PROGRESSBAR = True
READ_ONLY = True


class MdbLayer:
    """ Pretend we are a data provider """

    dirty = False
    doing_attr_update = False

    def __init__(self, mdb_path, mdb_table, mdb_columns='*', mdb_hide_columns = ''):
        """ Initialize the layer by reading a Access mdb file, creating a memory layer, and adding records to it

        :param mdb_path: Path to the database you wish to access.
        :type mdb_path: str

        :param mdb_table: Table name of the table to open.
        :type mdb_table: str

        :param mdb_columns: Comma separated list of columns to use for this layer. Defaults to all (*).
        :type mdb_columns: str

        :param mdb_hide_columns: Comma separated list of columns to hide for this layer.
            Use in combination with mdb_columns='*'
        :type mdb_hide_columns: str
        """

        self.mdb_path = mdb_path
        self.mdb_table = mdb_table
        self.mdb_columns = mdb_columns
        self.mdb_hide_columns = [x.strip() for x in mdb_hide_columns.split(",")]
        self.record_count = 0
        self.progress = object
        self.old_pk_values = {}
        self.read_only = READ_ONLY

        # connect to the database
        constr = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};FIL={MS Access};DBQ=" + self.mdb_path
        try:
            conn = pyodbc.connect(constr, timeout=3)
            self.cur = conn.cursor()
        except Exception as e:
            logger("Couldn't connect. Error: {}".format(e))
            return

        # determine primary key(s) if table
        table = self.cur.tables(table=self.mdb_table).fetchone()
        if table.table_type == 'TABLE':
            self.pk_cols = [row[8] for row in self.cur.statistics(self.mdb_table) if row[5] == 'PrimaryKey']
        elif table.table_type == 'VIEW':
            self.pk_cols = []
        else:
            logger("Database object type '{}' not supported".format(table.table_type))
            return

        # get record count
        try:
            self.cur.execute("SELECT COUNT(*) FROM {}".format(self.mdb_table))
            self.record_count = self.cur.fetchone()[0]
        except Exception as e:
            self.iface.messageBar().pushWarning("MDB Layer",
                "There's a problem with this table or query. Error: {}".format(e))
            return

        # get records from the table
        sql = "SELECT {} FROM {}".format(self.mdb_columns, self.mdb_table)
        self.cur.execute(sql)

        # create a dictionary with fieldname:type
        # QgsField only supports: String / Int / Double
        # falling back to string for: bytearray, bool, datetime.datetime, datetime.date, datetime.time, ?
        field_name_types = []
        field_type_map = {str: QVariant.String, unicode: QVariant.String,
                          int: QVariant.Int, float: QVariant.Double}

        # create a list with a QgsFields for every db column
        for column in self.cur.description:
            if column[0] not in self.mdb_hide_columns:
                if column[1] in field_type_map:
                    field_name_types.append(QgsField(column[0], field_type_map[column[1]]))
                else:
                    field_name_types.append(QgsField(column[0], QVariant.String))
                    self.read_only = True        # no reliably editing for other data types

        # create the layer, add columns
        self.lyr = QgsVectorLayer("None", 'mdb_' + mdb_table, 'memory')
        provider = self.lyr.dataProvider()
        provider.addAttributes(field_name_types)
        self.lyr.updateFields()

        # add the records
        self.add_records()

        # set read only or make connections/triggers
        # if there are no primary keys there is no way to edit
        if self.read_only or not self.pk_cols:
            self.lyr.setReadOnly()
        else:
            self.lyr.beforeCommitChanges.connect(self.before_commit)

        # add the layer the map
        QgsMapLayerRegistry.instance().addMapLayer(self.lyr)

    def add_records(self):
        """ Add records to the memory layer by stepping through the query result """

        self.setup_progressbar("Loading {} records from table {}..."
                               .format(self.record_count, self.lyr.name()),
                               self.record_count)

        provider = self.lyr.dataProvider()
        for i, row in enumerate(self.cur):
            feature = QgsFeature()
            feature.setGeometry(QgsGeometry())
            feature.setAttributes([flds for flds in row])
            provider.addFeatures([feature])
            self.update_progressbar(i)

        iface.messageBar().clearWidgets()
        iface.messageBar().pushMessage("Ready", "{} records added to {}".format(str(self.record_count), self.lyr.name())
                                       , level=QgsMessageBar.INFO)

    def before_commit(self):
        """" Just before a definitive commit (update to the memory layer) try
         updating the database"""
        # fixme: implement rollback on db fail

        # Update attribute values
        # ---------------------------------------------------------------------
        # changes: { fid: {pk1: value, pk2: value, etc}}
        changes = self.lyr.editBuffer().changedAttributeValues()
        field_names = [field.name() for field in self.lyr.pendingFields()]
        row_count = 0

        for fid, attributes in changes.iteritems():
            feature = self.lyr.dataProvider().getFeatures(QgsFeatureRequest(fid)).next()
            fields = [field_names[att_id] + " = (?)" for att_id in attributes]
            params = tuple(attributes.values())

            # assemble SQL query
            where_clause, params = self.get_where_clause(feature, params)
            sql = "UPDATE {} SET {}".format(self.mdb_table, ", ".join(fields))
            sql += where_clause

            logger(sql + " : " + str(params))
            self.cur.execute(sql, params)
            row_count += self.cur.rowcount

        self.cur.commit()
        logger("changed:  " + str(row_count))

        # Delete features: 'DELETE * FROM spoor WHERE pk = id
        # ---------------------------------------------------------------------
        row_count = 0
        fids = self.lyr.editBuffer().deletedFeatureIds()
        for feature in self.lyr.dataProvider().getFeatures(QgsFeatureRequest().setFilterFids(fids)):

            where_clause, params = self.get_where_clause(feature)

            # assemble SQL query
            sql = "DELETE * FROM {}".format(self.mdb_table)
            sql += where_clause

            logger(sql + " : " + str(params))
            self.cur.execute(sql, params)
            row_count += self.cur.rowcount

        self.cur.commit()
        logger("deleted:  " + str(row_count))

        # Add features
        # ---------------------------------------------------------------------
        # fixme: implement

    def get_where_clause(self, feature, params=()):
        """ Return where_clause with pk fields and updated params"""
        where_clause = []
        for pk in self.pk_cols:
            params += (feature[pk],)
            where_clause.append(pk + " = (?)")

        where_clause = " WHERE " + " AND ".join(where_clause)
        return where_clause, params

    def setup_progressbar(self, message, maximum):
        if not SHOW_PROGRESSBAR: return

        progress_message_bar = iface.messageBar().createMessage(message)
        self.progress = QProgressBar()
        self.progress.setMaximum(maximum)
        self.progress.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        progress_message_bar.layout().addWidget(self.progress)
        iface.messageBar().pushWidget(progress_message_bar, iface.messageBar().INFO)

    def update_progressbar(self, progress):
        if SHOW_PROGRESSBAR:
            self.progress.setValue(progress)