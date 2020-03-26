namespace Fortuna.Services.TemplateStorage.Interface.Dtos
{
    /// <summary>
    /// Placeholder-érték tároló absztrakt osztály.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class PlaceHolderNameValueDto<T> : PlaceHolderNameValueDtoBase
    {
        /// <summary>
        /// Placeholder értéke.
        /// </summary>
        public T Value { get; set; }
    }
}
