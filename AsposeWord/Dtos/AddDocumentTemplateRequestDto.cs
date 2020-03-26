namespace Fortuna.Services.TemplateStorage.Interface.Dtos
{
    /// <summary>
    /// Új dokumentum template hozzáadás kérés adatai.
    /// </summary>
    public class AddDocumentTemplateRequestDto
    {
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
    }
}
