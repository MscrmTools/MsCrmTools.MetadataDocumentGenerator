using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MsCrmTools.MetadataDocumentGenerator.Helper
{
    public static class Extensions
    {
        public static bool IsRollupDerivedColumn(this AttributeMetadata amd, IEnumerable<AttributeMetadata> allAmds)
        {
            string rootAttributeLogicalName = "";
            if (amd.LogicalName.EndsWith("_state"))
            {
                rootAttributeLogicalName = amd.LogicalName.Substring(0, amd.LogicalName.LastIndexOf("_state"));
                return allAmds.Any(a => a.LogicalName == rootAttributeLogicalName);
            }

            if (amd.LogicalName.EndsWith("_date"))
            {
                rootAttributeLogicalName = amd.LogicalName.Substring(0, amd.LogicalName.LastIndexOf("_date"));
                return allAmds.Any(a => a.LogicalName == rootAttributeLogicalName);
            }
            return false;
        }
    }
}