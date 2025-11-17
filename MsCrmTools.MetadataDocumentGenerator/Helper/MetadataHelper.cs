using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using XrmToolBox;

namespace MsCrmTools.MetadataDocumentGenerator.Helper
{
    /// <summary>
    /// Class for querying Crm Metadata
    /// </summary>
    internal class MetadataHelper
    {
        private const int SOLUTION_BATCH_SIZE = 50; // Maximum solutions to process in a single batch
        private const int ENTITY_METADATA_BATCH_SIZE = 100; // Maximum entity metadata IDs to retrieve in a single request

        public static List<EntityMetadata> GetEntities(List<Entity> solutions, IOrganizationService service)
        {
            var allObjectIds = new HashSet<Guid>(); // Use HashSet to avoid duplicates

            if (solutions.Count > 0)
            {
                // Process solutions in batches to avoid request size limits
                for (int i = 0; i < solutions.Count; i += SOLUTION_BATCH_SIZE)
                {
                    var batch = solutions.Skip(i).Take(SOLUTION_BATCH_SIZE).ToList();

                    var components = service.RetrieveMultiple(new QueryExpression("solutioncomponent")
                    {
                        ColumnSet = new ColumnSet("objectid"),
                        NoLock = true,
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("solutionid", ConditionOperator.In, batch.Select(s => s.Id).ToArray()),
                                new ConditionExpression("componenttype", ConditionOperator.Equal, 1)
                            }
                        }
                    }).Entities;

                    // Add unique object IDs to the set
                    foreach (var component in components)
                    {
                        allObjectIds.Add(component.GetAttributeValue<Guid>("objectid"));
                    }
                }
            }

            // Convert to list for batching
            var allEntityIds = allObjectIds.ToList();
            var allEntityMetadata = new List<EntityMetadata>();

            // Process entity metadata retrieval in batches to avoid large query issues
            for (int i = 0; i < allEntityIds.Count; i += ENTITY_METADATA_BATCH_SIZE)
            {
                var entityBatch = allEntityIds.Skip(i).Take(ENTITY_METADATA_BATCH_SIZE).ToList();

                EntityQueryExpression entityQueryExpression = new EntityQueryExpression
                {
                    Criteria = new MetadataFilterExpression(LogicalOperator.Or),
                    Properties = new MetadataPropertiesExpression
                    {
                        AllProperties = true
                    }
                };

                // Add metadata conditions for this batch
                entityBatch.ForEach(id =>
                {
                    entityQueryExpression.Criteria.Conditions.Add(
                        new MetadataConditionExpression("MetadataId", MetadataConditionOperator.Equals, id));
                });

                RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
                {
                    Query = entityQueryExpression,
                    ClientVersionStamp = null
                };

                var response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);
                allEntityMetadata.AddRange(response.EntityMetadata);
            }

            return allEntityMetadata;
        }

        /// <summary>
        /// Gets the list of entities metadata (only Entity Items)
        /// </summary>
        /// <returns>List of entities metadata</returns>
        public static List<EntityMetadata> RetrieveEntities(IOrganizationService oService)
        {
            var request = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
            };

            var response = (RetrieveAllEntitiesResponse)oService.Execute(request);

            return response.EntityMetadata.ToList();
        }

        /// <summary>
        /// Gets specified entity metadata (include attributes)
        /// </summary>
        /// <param name="logicalName">Logical name of the entity</param>
        /// <param name="oService">Crm organization service</param>
        /// <returns>Entity metadata</returns>
        public static EntityMetadata RetrieveEntity(string logicalName, IOrganizationService oService)
        {
            try
            {
                var request = new RetrieveEntityRequest
                {
                    LogicalName = logicalName,
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes,
                    RetrieveAsIfPublished = true
                };

                var response = (RetrieveEntityResponse)oService.Execute(request);

                return response.EntityMetadata;
            }
            catch (Exception error)
            {
                string errorMessage = CrmExceptionHelper.GetErrorMessage(error, false);
                throw new Exception("Error while retrieving entity: " + errorMessage);
            }
        }

        /// <summary>
        /// Retrieves main forms for the specified entity
        /// </summary>
        /// <param name="logicalName">Entity logical name</param>
        /// <param name="oService">Crm organization service</param>
        /// <returns>Document containing all forms definition</returns>
        public static IEnumerable<Entity> RetrieveEntityFormList(string logicalName, IOrganizationService oService)
        {
            var qe = new QueryExpression("systemform")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objecttypecode", ConditionOperator.Equal, logicalName),
                        new ConditionExpression("type", ConditionOperator.In, new[] {2,7}),
                    }
                }
            };

            try
            {
                return oService.RetrieveMultiple(qe).Entities;
            }
            catch
            {
                qe.Criteria.Conditions.RemoveAt(qe.Criteria.Conditions.Count - 1);
                return oService.RetrieveMultiple(qe).Entities;
            }
        }

        /// <summary>
        /// Retrieves main forms for the specified entity
        /// </summary>
        /// <param name="logicalName">Entity logical name</param>
        /// <param name="oService">Crm organization service</param>
        /// <returns>Document containing all forms definition</returns>
        public static List<XmlDocument> RetrieveEntityForms(string logicalName, List<Guid> formsIds, IOrganizationService oService)
        {
            var qe = new QueryExpression("systemform")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objecttypecode", ConditionOperator.Equal, logicalName),
                        new ConditionExpression("type", ConditionOperator.Equal, 2),
                    }
                }
            };

            if (formsIds.Count > 0)
            {
                qe.Criteria.AddCondition("formid", ConditionOperator.In, formsIds.Select(f => f.ToString()).ToArray());
            }

            var ec = oService.RetrieveMultiple(qe);

            var docs = new List<XmlDocument>();
            foreach (var form in ec.Entities)
            {
                var doc = new XmlDocument();
                doc.LoadXml(form.GetAttributeValue<string>("formxml"));
                docs.Add(doc);
            }

            return docs;
        }
    }
}