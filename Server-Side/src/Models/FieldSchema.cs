using System.Collections.Generic;

namespace EJ2APIServices.Models
{
    /// <summary>
    /// Merge field schema served by GET /api/schema.
    /// The Angular Merge Field Explorer is built entirely from this response,
    /// so the front end has no hardcoded field definitions.
    /// </summary>
    public class FieldSchema
    {
        /// <summary>Case-level merge fields (one value per case).</summary>
        public Dictionary<string, string> Case { get; set; }

        /// <summary>Party-level merge fields (repeat for every party in a case).</summary>
        public Dictionary<string, string> Parties { get; set; }
    }
}