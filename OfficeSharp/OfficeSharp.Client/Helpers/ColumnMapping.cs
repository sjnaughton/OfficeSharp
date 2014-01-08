using System;
using System.Linq;
using System.Collections.Generic;

namespace OfficeSharp
{
    public class ColumnMapping
    {
        public ColumnMapping(string OfficeColumnName, List<FieldDefinition> TableFields)
        {
            this.OfficeColumn = OfficeColumnName;
            this.TableFields = TableFields;

            //Try match up the column name to the table property choice
            TableField = TableFields.Where(f => f.Name == OfficeColumnName || f.DisplayName == OfficeColumnName).FirstOrDefault();
        }
        public ColumnMapping(string OfficeColumnName, string EntityFieldName)
        {
            this.OfficeColumn = OfficeColumnName;
            this.TableField = new FieldDefinition(EntityFieldName);
        }
        public string OfficeColumn { get; set; }
        public FieldDefinition TableField { get; set; }
        public List<FieldDefinition> TableFields { get; set; }
    }
}
