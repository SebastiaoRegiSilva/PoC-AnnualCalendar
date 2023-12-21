using System.Xml.Serialization;

/// <summary>
/// Capítulo de cada livro.
/// </summary>
[Serializable]
    public class Capitulo
    {
        [XmlElement("v")]
        public List<Versiculo>? Versiculos { get; set; }
        
        [XmlElement("n")]
        public int NumeroDoCapitulo { get; set; }
        
        [XmlElement("text")]
        public string? Text { get; set; }

    public Capitulo(List<Versiculo>? versiculos, int numeroDoCapitulo, string? text)
    {
        Versiculos = versiculos;
        NumeroDoCapitulo = numeroDoCapitulo;
        Text = text;
    }

    public Capitulo()
    {
    }
}