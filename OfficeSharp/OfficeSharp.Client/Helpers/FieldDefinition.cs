using System;
using System.Collections.Generic;
using Microsoft.LightSwitch.Model;

namespace OfficeSharp
{
    public class FieldDefinition
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string TypeName { get; set; }
        public bool IsNullable { get; set; }
        public IEntityType EntityType { get; set; }

        public FieldDefinition()
        {
        }

        public FieldDefinition(string name, string displayName, string typeName, bool isNullable)
        {
            this.Name = name;
            this.DisplayName = displayName;
            this.TypeName = typeName;
            this.IsNullable = isNullable;
        }

        public FieldDefinition(string name)
        {
            this.Name = name;
            this.DisplayName = name;
        }
    }
}
