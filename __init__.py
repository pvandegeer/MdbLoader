# -*- coding: utf-8 -*-
"""
/***************************************************************************
 MdbLoader
                                 A QGIS plugin
 This plugin loads a MS Access table for use in QGIS
                             -------------------
        begin                : 2017-09-22
        copyright            : (C) 2017 by P. van de Geer
        email                : pvandegeer@gmail.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load MdbLoader class from file MdbLoader.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .mdb_loader import MdbLoader
    return MdbLoader(iface)
