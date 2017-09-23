This plugin loads a Microsoft Access table for use in QGIS.

The underlying "pseudo mdb provider" simulates a data provider using a memory layer.
This concept is copied from the Pseudo CSV Provider which can be found here:
http://github.com/g-sherman/pseudo_csv_provider
The concept is the same, the actual implementation is different.

Like the original, the provider will:
* create a memory layer from a table in a MDB file
* create fields in the layer based on the different datatype found in the table
* write changes back to the database table using the primary keys (experimental, read-only by default)
* not support geometries, but can easily linked to another layer

A few things about this implementation:
* It is an example, not a robust implementation
* It lacks proper error handling
* It could be extended to support geometry types