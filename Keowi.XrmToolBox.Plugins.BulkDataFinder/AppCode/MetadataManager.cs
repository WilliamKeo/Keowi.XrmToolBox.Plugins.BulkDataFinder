using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Keowi.XrmToolBox.Plugins.BulkDataFinder.AppCode
{
    public class MetadataManager
    {
        private readonly IOrganizationService Service;

        public MetadataManager(IOrganizationService service)
        {
            Service = service;
        }

        public List<EntityMetadata> GetEntitiesMetadata()
        {
            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            metaDataRequest.EntityFilters = EntityFilters.Entity;
            RetrieveAllEntitiesResponse metaDataResponse = (RetrieveAllEntitiesResponse)Service.Execute(metaDataRequest);

            var entitiesResponse = metaDataResponse.EntityMetadata;

            var filteredEntities = entitiesResponse.Where(x => x.IsPrivate == false
                        && x.IsIntersect == false).OrderBy(x => x.LogicalName).ToList();

            return filteredEntities;
        }

        public Tuple<string, List<string>, List<AttributeMetadata>> GetEntityPrimaryAndTextAttributes(string entityLogicalName)
        {
            var request = new RetrieveEntityRequest
            {
                RetrieveAsIfPublished = true,
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entityLogicalName
            };

            var result = (RetrieveEntityResponse)Service.Execute(request);
            //Remove internal attributes.
            var filteredAttributes = result.EntityMetadata.Attributes.Where(a =>
                a.AttributeOf == null);
            var attributesMetadata = filteredAttributes;
            //Remove non string attributes or primary key.
            filteredAttributes = filteredAttributes.Where(a => a.AttributeType == AttributeTypeCode.Uniqueidentifier ||
                a.AttributeType == AttributeTypeCode.String || a.AttributeType == AttributeTypeCode.Memo);

            return new Tuple<string, List<string>, List<AttributeMetadata>>(
                result.EntityMetadata.PrimaryNameAttribute,
                filteredAttributes.Select(x => x.LogicalName).OrderBy(x => x).ToList(),
                attributesMetadata.ToList());
        }

        public EntityCollection GetSavedQueries(int entityCode)
        {
            return Service.RetrieveMultiple(new QueryExpression("savedquery")
            {
                NoLock = true,
                ColumnSet = new ColumnSet("savedqueryid", "name", "fetchxml"),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("returnedtypecode", ConditionOperator.Equal, entityCode),
                        new ConditionExpression("isquickfindquery", ConditionOperator.Equal, false),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                    }
                }
            });
        }

        public EntityCollection GetUserQueries(int entityCode)
        {
            return Service.RetrieveMultiple(new QueryExpression("userquery")
            {
                NoLock = true,
                ColumnSet = new ColumnSet("userqueryid", "name", "fetchxml"),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("returnedtypecode", ConditionOperator.Equal, entityCode),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("ownerid", ConditionOperator.EqualUserId)
                    }
                }
            });
        }

        public OptionMetadata[] GetOptionSetMetadata(string entityName, string attributeName)
        {
            var retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };

            var retrieveAttributeResponse = (RetrieveAttributeResponse)Service.Execute(retrieveAttributeRequest);
            var retrievedPicklistAttributeMetadata =
                (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();

            return optionList;
        }

        public OptionMetadata[] GetMultiSelectOptionSetMetadata(string entityName, string attributeName)
        {
            var retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };

            var retrieveAttributeResponse = (RetrieveAttributeResponse)Service.Execute(retrieveAttributeRequest);
            var retrievedMultiSelectPicklistAttributeMetadata =
                (MultiSelectPicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            OptionMetadata[] optionList = retrievedMultiSelectPicklistAttributeMetadata.OptionSet.Options.ToArray();

            return optionList;
        }

        public string GetOptionSetText(OptionMetadata[] optionSet, int optionSetValue)
        {
            if (optionSet.Length > 0)
            {
                foreach (OptionMetadata optionMetadata in optionSet)
                {
                    if (optionMetadata.Value == optionSetValue)
                    {
                        return optionMetadata.Label.UserLocalizedLabel.Label;
                    }
                }
            }
            return null;
        }
    }
}