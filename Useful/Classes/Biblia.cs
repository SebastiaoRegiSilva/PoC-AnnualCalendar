using System.Xml.Serialization;

/// <summary>
/// Bíblia Sagrada.
/// </summary>
[XmlRoot("bible")]
    public class Biblia
    {
        [XmlElement("book")]
        public Livro? Livros { get; set; }

    /// <summary>
    /// Construtor de objeto com parâmetro.
    /// </summary>
    /// <param name="livro"></param>
    public Biblia(Livro livro)
    {
        Livros = livro;
    }

    public Biblia()
    {
    }
}