using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DocxTranslator
{
    [XmlRoot(ElementName = "TranslateArrayRequest", IsNullable = false)]
    public class TranslateArrayRequestElement
    {
        [XmlElement(ElementName = "AppId", IsNullable = false)]
        public string AppId { get; set; }

        [XmlElement(ElementName = "From", IsNullable = false)]
        public string From { get; set; }

        [XmlElement(ElementName = "Options", IsNullable = false)]
        public OptionsElement Options { get; set; }

        [XmlArray(ElementName = "Texts", IsNullable = false)]
        [XmlArrayItem(ElementName = "string", Type = typeof(string), IsNullable = false, Namespace = "http://schemas.microsoft.com/2003/10/Serialization/Arrays")]
        public string[] Texts { get; set; }

        [XmlElement(ElementName = "To", IsNullable = false)]
        public string To { get; set; }
    }

    public class OptionsElement
    {
        [XmlElement(ElementName = "Category", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string Category { get; set; }

        [XmlElement(ElementName = "ContentType", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string ContentType { get; set; }

        [XmlElement(ElementName = "ReservedFlags", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string ReservedFlags { get; set; }

        [XmlElement(ElementName = "State", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string State { get; set; }

        [XmlElement(ElementName = "Uri", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string Uri { get; set; }

        [XmlElement(ElementName = "User", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string User { get; set; }

        [XmlElement(ElementName = "ProfanityAction", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
        public string ProfanityAction { get; set; }
    }

    [XmlRoot(ElementName = "ArrayOfTranslateArrayResponse", IsNullable = false, Namespace = "http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2")]
    public class ArrayOfTranslateArrayResponseElement
    {
        [XmlElement(ElementName = "TranslateArrayResponse", IsNullable = false)]
        public TranslateArrayResponseElement[] TranslateArrayResponse { get; set; }
    }

    public class TranslateArrayResponseElement
    {
        [XmlElement(ElementName = "From", IsNullable = false)]
        public string From { get; set; }

        [XmlArray(ElementName = "OriginalTextSentenceLengths", IsNullable = false)]
        [XmlArrayItem(ElementName = "int", Type = typeof(int), IsNullable = false, Namespace = "http://schemas.microsoft.com/2003/10/Serialization/Arrays")]
        public int[] OriginalTextSentenceLengths { get; set; }

        [XmlElement(ElementName = "State", IsNullable = false)]
        public string State { get; set; }

        [XmlElement(ElementName = "TranslatedText", IsNullable = false)]
        public string TranslatedText { get; set; }

        [XmlArray(ElementName = "TranslatedTextSentenceLengths", IsNullable = false)]
        [XmlArrayItem(ElementName = "int", Type = typeof(int), IsNullable = false, Namespace = "http://schemas.microsoft.com/2003/10/Serialization/Arrays")]
        public int[] TranslatedTextSentenceLengths { get; set; }
    }
}
