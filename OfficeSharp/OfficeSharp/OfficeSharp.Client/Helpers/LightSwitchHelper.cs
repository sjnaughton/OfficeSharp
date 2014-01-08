using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.LightSwitch;
using Microsoft.LightSwitch.Client;
using Microsoft.LightSwitch.Details;
using Microsoft.LightSwitch.Model;
using Microsoft.LightSwitch.Presentation;
using Microsoft.LightSwitch.Presentation.Extensions;

namespace OfficeSharp
{
	public static class LightSwitchHelper
	{

		public static string GetValue(this IEntityObject Entity, string FieldName)
		{
			string sValue = "";
			try
			{
				if ((Entity.Details.Properties[FieldName].Value != null))
				{
					if (Entity.Details.Properties[FieldName].PropertyType.ToString() == "System.Byte[]")
						sValue = "(Image)";
					else if (Entity.Details.Properties[FieldName].PropertyType.ToString().StartsWith("Microsoft.LightSwitch.Framework.EntityCollection"))
						sValue = "(Collection)";
					else
						sValue = Entity.Details.Properties[FieldName].Value.ToString();
				}
			}
			catch
			{
			}
			return sValue;
		}


		public static string GetValue(string Item)
		{
			string sValue = "";
			try
			{
				if ((Item != null))
				{
					if (Item == "System.Byte[]")
						sValue = "(Image)";
					else if (Item.StartsWith("Microsoft.LightSwitch.Framework.EntityCollection"))
						sValue = "(Collection)";
					else
						sValue = Item;
				}
			}
			catch
			{
			}
			return sValue;
		}

		public static bool ValidData(string value, int row, ColumnMapping mapping, List<string> errorList)
		{
			dynamic convertedValue = null;
			bool isValid = TryConvertValue(mapping.TableField.TypeName, value, ref convertedValue);
			if (isValid == false)
				errorList.Add(String.Format("Column:{0} Row:{1} Cannot convert value({2}) to {3} for '{4}'", mapping.OfficeColumn, row, value, mapping.TableField.TypeName, mapping.TableField.DisplayName));

			return isValid;
		}

		public static bool ValidData(string value, int row, ColumnMapping mapping, IVisualCollection collection, Dictionary<string, IEntityObject> navProperties, List<string> errorList)
		{
			bool bValid = false;
			IEntityType targetEntityType = mapping.TableField.EntityType;

			//Need to grab the entity set and check number of results we get
			IApplicationDefinition appModel = collection.Screen.Details.Application.Details.GetModel();
			IEntityContainerDefinition entityContainerDefinition = (from ecd in appModel.GlobalItems.OfType<IEntityContainerDefinition>()
																	where ecd.EntitySets.Any(es => object.ReferenceEquals(es.EntityType, targetEntityType))
																	select ecd).FirstOrDefault();

			if (entityContainerDefinition == null)
				throw new Exception("Could not find an entity container representing the entity type: " + targetEntityType.Name);

			IEntitySetDefinition entitySetDefinition = (from es in entityContainerDefinition.EntitySets
														where object.ReferenceEquals(es.EntityType, targetEntityType)
														select es).First();

			var dataService = (IDataService)collection.Screen.Details.DataWorkspace.Details.Properties[entityContainerDefinition.Name].Value;
			var entitySet = (IEntitySet)dataService.Details.Properties[entitySetDefinition.Name].Value;
			var dsQuery = entitySet.GetQuery();

			//Search for the matching entity for the relationship IEnumerable<IEntityObject>
			var results = SearchEntityMethodInfo().MakeGenericMethod(dsQuery.ElementType).Invoke(null,
				new object[] 
				{
					dsQuery,
					value,
					targetEntityType
				}) as IEnumerable<IEntityObject>;

			int searchCount = results.Count();
			if (searchCount == 0)
			{
				bValid = false;
				errorList.Add(String.Format("Column:{0} Row:{1} Cannot find a matching '{2}' for '{3}'", mapping.OfficeColumn, row, mapping.TableField.DisplayName, value));
			}
			else if (searchCount > 1)
			{
				bValid = true;
				errorList.Add(String.Format("Column:{0} Row:{1} Multiple matching '{2}' for '{3}'.  Will select first match.", mapping.OfficeColumn, row, mapping.TableField.DisplayName, value));
				navProperties[String.Format("{0}_{1}", mapping.TableField.Name, value)] = results.FirstOrDefault();
			}
			else
			{
				bValid = true;
				navProperties[String.Format("{0}_{1}", mapping.TableField.Name, value)] = results.FirstOrDefault();
			}
			return bValid;
		}

		public static IEntityPropertyDefinition GetSummaryProperty(IEntityType entityType)
		{
			//Return the specified property if one is specified
			ISummaryPropertyAttribute attribute = entityType.Attributes.OfType<ISummaryPropertyAttribute>().FirstOrDefault();
			if (attribute != null && attribute.Property != null)
				return attribute.Property;

			//If none is specified, try to infer one
			IEnumerable<IEntityPropertyDefinition> properties = entityType.Properties.Where(p => (!(p is INavigationPropertyDefinitionBase)) && (!p.PropertyType.Name.Contains("Binary")));
			IEntityPropertyDefinition stringProperty = properties.FirstOrDefault(p => p.PropertyType.Name.Contains("String"));
			if (stringProperty == null)
				return properties.FirstOrDefault();
			else
				return stringProperty;
		}

		// Returns a FieldDefinition object given a collection and a field name. Does not work for collection fields or entity fields
		public static FieldDefinition GetFieldDefinition(this IVisualCollection collection, string FieldName)
		{
			var entityType = (IEntityType)collection.Details.GetModel().ElementType;
			bool isNullable = false;
			FieldDefinition fd = null;

			foreach (IEntityPropertyDefinition p in entityType.Properties)
			{
				//Ignore hidden fields and computed field
				if (p.Attributes.Where(a => a.Class.Name == "Computed").FirstOrDefault() == null)
				{
					if (!(p.PropertyType is ISequenceType))
					{
						//ignore collections and entities
						if (p.Name == FieldName)
						{
							fd = new FieldDefinition();
							fd.Name = p.Name;
							fd.DisplayName = p.Name;
							fd.TypeName = GetPropertyType(p.PropertyType, ref isNullable);
							fd.IsNullable = isNullable;
							if (fd.TypeName == "Entity")
							{
								fd.EntityType = (IEntityType)p.PropertyType;
							}
							break; // TODO: might not be correct. Was : Exit For
						}
					}
				}
			}

			return fd;
		}

		public static string GetPropertyType(this IDataType p, ref bool isNullable)
		{
			string typeName = "";
			if ((p) is ISemanticType)
			{
				typeName = ((ISemanticType)p).UnderlyingType.Name;
				isNullable = false;
			}
			else if ((p) is INullableType && (((INullableType)p).UnderlyingType) is ISemanticType)
			{
				typeName = ((ISemanticType)((INullableType)p).UnderlyingType).UnderlyingType.Name;
				isNullable = true;
			}
			else if ((p) is INullableType)
			{
				typeName = ((INullableType)p).UnderlyingType.Name;
				isNullable = true;
			}
			else if ((p) is ISimpleType)
			{
				typeName = ((ISimpleType)p).Name;
				isNullable = false;
			}
			else if ((p) is IEntityType)
			{
				typeName = "Entity";
				isNullable = true;
				//fd.EntityType = DirectCast(p.PropertyType, IEntityType)
			}
			else
			{
				throw new NotSupportedException("Could not determine the property type");
			}
			return typeName;
		}

		public static bool IsComputed(IEntityPropertyDefinition entityProperty)
		{
			return entityProperty.Attributes.OfType<IComputedAttribute>().Any();
		}

		public static bool TryConvertValue(string propertyType, string value, ref object convertedValue)
		{
			bool canConvert = false;

			String convertedStringTry = null;
			Boolean convertedBooleanTry;
			DateTime convertedDateTimeTry;
			Decimal convertedDecimalTry;
			Double convertedDoubleTry;
			Int32 convertedInt32Try;
			Int16 convertedInt16Try;
			Int64 convertedInt64Try;

			switch (propertyType)
			{
				case "Binary":
					canConvert = false;
					break;
				case "Boolean":
					canConvert = Boolean.TryParse(value, out convertedBooleanTry);
					break;
				case "DateTime":
					canConvert = DateTime.TryParse(value, out convertedDateTimeTry);
					break;
				case "Decimal":
					canConvert = Decimal.TryParse(value, out convertedDecimalTry);
					break;
				case "Double":
					canConvert = Double.TryParse(value, out convertedDoubleTry);
					break;
				case "int":
					canConvert = Int32.TryParse(value, out convertedInt32Try);
					break;
				case "Int16":
					canConvert = Int16.TryParse(value, out convertedInt16Try);
					break;
				case "Int64":
					canConvert = Int64.TryParse(value, out convertedInt64Try);
					break;
				case "String":
					canConvert = true;
					convertedStringTry = value;
					break;
				default:
					throw new NotSupportedException();
			}

			if (canConvert)
				convertedValue = convertedStringTry;

			return canConvert;
		}

		#region "Generic method to search an entity"
		private static MethodInfo _SearchEntityMethodInfo;
		public static MethodInfo SearchEntityMethodInfo()
		{
			if (_SearchEntityMethodInfo == null)
			{
				_SearchEntityMethodInfo = typeof(LightSwitchHelper).GetMethod("SearchEntity", BindingFlags.Static | BindingFlags.NonPublic);
				Debug.Assert(_SearchEntityMethodInfo != null);
			}
			return _SearchEntityMethodInfo;
		}

		//private static IEnumerable<IEntityObject> SearchEntity<T>(IDataServiceQueryable<T> query, string value, IEntityType entityType) where T : IEntityObject
		//{
		//    //LOGIC
		//    //Search PK field (if only one)
		//    //Search summary field (if not computed)
		//    //Do a generic search
		//        Boolean isNullable=false;

		//    if (entityType.KeyProperties.Count() == 1) 
		//    {
		//        IKeyPropertyDefinition entityKey = entityType.KeyProperties.First();
		//        string propertyType = GetPropertyType(entityKey.PropertyType, ref isNullable);

		//        dynamic convertedValue = null;
		//        if (TryConvertValue(propertyType, value, ref convertedValue)) 
		//        {
		//            ParameterExpression pe = Expression.Parameter(query.ElementType, "entity");
		//            Expression wherePredicate = Expression.Equal(Expression.Property(pe, entityKey.Name), Expression.Constant(convertedValue));
		//            var whereExpression = Expression.Lambda(wherePredicate, pe);

		//            IEnumerable<IEntityObject> keyResults = query.Where(whereExpression).Execute.Cast<IEntityObject>();

		//            if (keyResults.Count > 0) 
		//                return keyResults;
		//        }
		//    }

		//    IEntityPropertyDefinition summaryProperty = GetSummaryProperty(entityType);
		//    if (summaryProperty != null && !IsComputed(summaryProperty)) 
		//    {
		//        string propertyType = GetPropertyType(summaryProperty.PropertyType, ref isNullable);

		//        dynamic convertedValue = null;
		//        if (TryConvertValue(propertyType, value, ref convertedValue)) 
		//        {
		//            ParameterExpression pe = Expression.Parameter(typeof(T), "entity");
		//            Expression wherePredicate = Expression.Equal(Expression.Property(pe, summaryProperty.Name), Expression.Constant(convertedValue));
		//            var whereExpression = Expression.Lambda(wherePredicate, pe);

		//            IEnumerable<IEntityObject> summaryResults = query.Where(whereExpression).Execute.Cast<IEntityObject>();

		//            if (summaryResults.Count > 0) 
		//                return summaryResults;
		//        }
		//    }


		//    return query.Search({ value }).Execute().Cast<IEntityObject>();
		//}
		#endregion
	}
}
