using System;
using System.Collections.Generic;

namespace Fortuna.Services.TemplateStorage.Interface.Dtos
{
    /// <summary>
    /// PDF generáláshoz szükséges adatok.
    /// </summary>
    public class PdfGenerationRequestDto
    {
        /// <summary>
        /// Template azonosítója.
        /// </summary>
        public Guid TemplateId { get; set; }

        /// <summary>
        /// Egyszerű kitöltési adatok név-érték listája.
        /// </summary>
        public List<SimplePlaceHolderNameValueDto> SimplePlaceholderNameValueList { get; set; }

        /// <summary>
        /// List kitöltési adatok név-érték listája.
        /// </summary>
        public List<ListPlaceHolderNameValueDto> ListPlaceholderNameValueList { get; set; }
    }
}