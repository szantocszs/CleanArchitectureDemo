using System;

namespace Fortuna.Services.TemplateStorage.Interface.Dtos
{
    /// <summary>
    /// Dukumentum template adatait tároló dto.
    /// </summary>
    public class DocumentTemplateDto
    {
        /// <summary>
        /// Template aznosító.
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// Template neve.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Template leírása.
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// Template.
        /// </summary>
        public byte[] Template { get; set; }

        /// <summary>
        /// Template létrehozásának ideje.
        /// </summary>
        public DateTime CreateTimeUtc { get; set; }
    }
}