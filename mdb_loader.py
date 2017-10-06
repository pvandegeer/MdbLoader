# -*- coding: utf-8 -*-
"""
/***************************************************************************
 mdb_loader
                                 A QGIS plugin
 This plugin loads a MS Access table for use in QGIS
                              -------------------
        begin                : 2017-09-22
        git sha              : $Format:%H$
        copyright            : (C) 2017 by P. van de Geer
        email                : pvandegeer@gmail.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from contextlib import contextmanager
from mdb_layer import MdbLayer
from PyQt4.QtCore import Qt, QSettings, QTranslator, qVersion, QCoreApplication
from PyQt4.QtGui import QApplication, QCursor, QAction, QIcon, QFileDialog
from qgis.core import QgsMessageLog, QgsProject

# Initialize Qt resources from file resources.py
import resources, pyodbc
# Import the code for the dialog
from mdb_loader_select_table import MdbLoaderSelectTable
import os.path

logger = lambda msg: QgsMessageLog.logMessage(msg, 'Mdb Loader', 1)

class MdbLoader:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'MdbLoader_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)


        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&Mdb Loader')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'MdbLoader')
        self.toolbar.setObjectName(u'MdbLoader')

    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('MdbLoader', message)

    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):
        """Add a toolbar icon to the toolbar.

        :param icon_path: Path to the icon for this action. Can be a resource
            path (e.g. ':/plugins/foo/bar.png') or a normal file system path.
        :type icon_path: str

        :param text: Text that should be shown in menu items for this action.
        :type text: str

        :param callback: Function to be called when the action is triggered.
        :type callback: function

        :param enabled_flag: A flag indicating if the action should be enabled
            by default. Defaults to True.
        :type enabled_flag: bool

        :param add_to_menu: Flag indicating whether the action should also
            be added to the menu. Defaults to True.
        :type add_to_menu: bool

        :param add_to_toolbar: Flag indicating whether the action should also
            be added to the toolbar. Defaults to True.
        :type add_to_toolbar: bool

        :param status_tip: Optional text to show in a popup when mouse pointer
            hovers over the action.
        :type status_tip: str

        :param parent: Parent widget for the new action. Defaults None.
        :type parent: QWidget

        :param whats_this: Optional text to show in the status bar when the
            mouse pointer hovers over the action.

        :returns: The action that was created. Note that the action is also
            added to self.actions list.
        :rtype: QAction
        """

        # Create the dialog (after translation) and keep reference
        self.dlg = MdbLoaderSelectTable()

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToDatabaseMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        icon_path = ':/plugins/mdb_loader/mdb_database.png'
        self.add_action(
            icon_path,
            text=self.tr(u'Open MS Access Table'),
            callback=self.run,
            parent=self.iface.mainWindow())


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginDatabaseMenu(
                self.tr(u'&Mdb Loader'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar


    def run(self):
        """Run method that performs all the real work"""
        mdb_file = QFileDialog.getOpenFileName(None, "Select database file",
                get_default_path(), 'Ms Access Database (*.mdb *.accdb)')
        if not mdb_file: return

        # store path; check if file exists
        set_default_path(mdb_file)
        if not os.path.isfile(mdb_file):
            self.iface.messageBar().pushError("MDB Loader", "File not found")
            return

         # connect to the database, get tables
        constr = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};FIL={MS Access};DBQ=" + mdb_file
        conn = pyodbc.connect(constr)
        cur = conn.cursor()

        with wait_cursor():
            self.dlg.listWidget.clear()
            for row in cur.tables():
                # todo: allow (readonly) queries?
                if row.table_type == 'TABLE':
                    #logger(row.table_name)
                    self.dlg.listWidget.addItem(row.table_name)

            conn.close()

        # return to QGis if there are no tables
        if self.dlg.listWidget.count() < 1:
            self.iface.messageBar().pushWarning("MDB Loader", "No tables were found")
            return

        # show the dialog
        self.dlg.listWidget.item(0).setSelected(True)
        self.dlg.show()

        # run the dialog event loop / see if OK was pressed
        result = self.dlg.exec_()
        if result:
            selected_table = self.dlg.listWidget.selectedItems()[0].text()
            self.mdblayer = MdbLayer(mdb_file, selected_table, mdb_hide_columns = 's_GUID, MAPINFO_ID')


def set_default_path(path):
    """Set the default path"""
    path = os.path.dirname(path)
    QSettings().setValue("mdb_loader/default_path", path)
    QgsProject.instance().writeEntry("mdb_loader", "default_path", path)
    return

def get_default_path():
    """Choose a sane default path"""
    path = QSettings().value("mdb_loader/default_path", os.environ['HOME'])
    path = QgsProject.instance().readEntry("mdb_loader", "default_path", path)[0]
    return path

@contextmanager
def wait_cursor():
    try:
        QApplication.setOverrideCursor(QCursor(Qt.WaitCursor))
        yield
    finally:
        QApplication.restoreOverrideCursor()