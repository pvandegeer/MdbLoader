import pyodbc, datetime
from PyQt4.QtGui import QProgressBar
from PyQt4.QtCore import QVariant, Qt
from qgis.utils import iface, QgsMessageBar
from qgis.core import (QgsVectorLayer, QgsFeature, QgsGeometry, QgsPoint, QgsField,
                       QgsMapLayerRegistry, QgsFeatureRequest, QgsMessageLog)


logger = lambda msg: QgsMessageLog.logMessage(msg, 'Mdb Provider Example', 1)

SHOW_PROGRESSBAR = True
READ_ONLY = True


class MdbLayer:
    """ Pretend we are a data provider """

    dirty = False
    doing_attr_update = False

    def __init__(self, mdb_path, mdb_table, mdb_columns='*', mdb_hide_columns = ''):
        """ Initialize the layer by reading a Access mdb file, creating a memory layer, and adding records to it """

        self.mdb_path = mdb_path
        self.mdb_table = mdb_table
        self.mdb_columns = mdb_columns
        self.mdb_hide_columns = [x.strip() for x in mdb_hide_columns.split(",")]
        self.record_count = 0
        self.old_pk_values = {}

        # connect to the database
        constr = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};FIL={MS Access};DBQ=" + self.mdb_path
        conn = pyodbc.connect(constr)
        self.cur = conn.cursor()

        # determine primary key(s)
        # fixme: make read only or fail if no pks available
        # todo: use primaryKeys(table, catalog=None, schema=None) ??
        self.pk_cols = [row[8] for row in self.cur.statistics(self.mdb_table) if row[5]=='PrimaryKey']

        # get record count
        self.record_count = self.cur.execute("SELECT COUNT(*) FROM " + self.mdb_table).fetchone()[0]

        # get records from the table
        # fixme: use parameters
        sql = "SELECT " + self.mdb_columns + " FROM " + self.mdb_table
        self.cur.execute(sql)
        # logger("sql:  " + sql)

        # create a dictionary with fieldname:type
        # fixme: QgsField: Field variant type, currently supported: String / Int / Double
        field_name_types = []
        field_type_map = {str: QVariant.String, unicode: QVariant.String,
                          int: QVariant.Int, float: QVariant.Double,
                          bytearray: QVariant.String,  bool: QVariant.String,
                          datetime.datetime: QVariant.String, datetime.date: QVariant.String,
                          datetime.time: QVariant.String}

        # create a list with a QgsFields for every db column
        for column in self.cur.description:
            # fixme: check for unhandled type
            # logger(str(column[0]) + ": " + str(column[1]))
            if column[0] not in self.mdb_hide_columns:
                field_name_types.append(QgsField(column[0], field_type_map[column[1]]))

        # create the layer, add columns
        self.lyr = QgsVectorLayer("None", 'mdb_' + mdb_table, 'memory')
        provider = self.lyr.dataProvider()
        provider.addAttributes(field_name_types)
        self.lyr.updateFields()

        # add the records
        self.add_records()

        # set read only or make connections/triggers
        if READ_ONLY:
            self.lyr.setReadOnly()
        else:
            self.lyr.beforeCommitChanges.connect(self.before_commit)

        # add the layer the map
        QgsMapLayerRegistry.instance().addMapLayer(self.lyr)

    def add_records(self):
        """ Add records to the memory layer by stepping through the query result """

        # set up a progressbar
        if SHOW_PROGRESSBAR:
            progress_message_bar = iface.messageBar().createMessage(
                "Loading {} records from table {}...".format(self.record_count, self.lyr.name()))
            progress = QProgressBar()
            progress.setMaximum(self.record_count)
            progress.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            progress_message_bar.layout().addWidget(progress)
            iface.messageBar().pushWidget(progress_message_bar, iface.messageBar().INFO)

        provider = self.lyr.dataProvider()
        for i, row in enumerate(self.cur):
            feature = QgsFeature()
            feature.setGeometry(QgsGeometry())
            feature.setAttributes([flds for flds in row])
            provider.addFeatures([feature])
            if SHOW_PROGRESSBAR: progress.setValue(i)

        iface.messageBar().clearWidgets()
        iface.messageBar().pushMessage("Ready", "{} records added to {}".format(str(self.record_count), self.lyr.name())
                                       , level=QgsMessageBar.INFO)

    def before_commit(self):
        """" Just before a definitive commit (update to the memory layer) try
         updating the database"""
        # fixme: implement rollback on db fail

        # Update attribute values
        # changes: { fid: {pk1: value, pk2: value, etc}}
        changes = self.lyr.editBuffer().changedAttributeValues()
        field_names = [field.name() for field in self.lyr.pendingFields()]
        row_count = 0

        for fid, attributes in changes.iteritems():
            feature = self.lyr.dataProvider().getFeatures(QgsFeatureRequest(fid)).next()
            fields = [field_names[att_id] + " = (?)" for att_id in attributes]
            params = tuple(attributes.values())

            where_clause, params = self.get_where_clause(feature, params)

            # assemble SQL query
            sql = "UPDATE " + self.mdb_table
            sql += " SET " + ", ".join(fields)
            sql += where_clause

            logger(sql + " : " + str(params))
            self.cur.execute(sql, params)
            row_count += self.cur.rowcount

        self.cur.commit()
        logger("changed:  " + str(row_count))

        # Delete features: 'DELETE * FROM spoor WHERE pk = id
        row_count = 0
        fids = self.lyr.editBuffer().deletedFeatureIds()
        for feature in self.lyr.dataProvider().getFeatures(QgsFeatureRequest().setFilterFids(fids)):

            where_clause, params = self.get_where_clause(feature)

            # assemble SQL query
            sql = "DELETE * FROM " + self.mdb_table
            sql += where_clause

            logger(sql + " : " + str(params))
            self.cur.execute(sql, params)
            row_count += self.cur.rowcount

        self.cur.commit()
        logger("deleted:  " + str(row_count))

        # Add features
        # fixme: implement

    def get_where_clause(self, feature, params=()):
        """ Return where_clause with pk fields and updated params"""
        where_clause = []
        for pk in self.pk_cols:
            params += (feature[pk],)
            where_clause.append(pk + " = (?)")

        where_clause = " WHERE " + " AND ".join(where_clause)
        return where_clause, params
