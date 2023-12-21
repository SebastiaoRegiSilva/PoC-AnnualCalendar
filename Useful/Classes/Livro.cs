using System.Xml.Serialization;

/// <summary>
/// Livro da Bíblia
/// </summary>
[Serializable]
    public class Livro
    {
        [XmlElement("c")]
        public Capitulo? Capitulos { get; set; }
        
        [XmlElement("name")]
        public string? Name { get; set; }
        
        [XmlElement("abbrev")]
        public string? Abbre { get; set; }
        
        [XmlElement("chapters")]
        public int QtdCapitulos { get; set; }
        
        [XmlElement("text")]
        public string? Text { get; set; }

    /// <summary>
    /// Construtor do objeto com parâmetros.
    /// </summary>
    /// <param name="capitulos"></param>
    /// <param name="name"></param>
    /// <param name="abbrev"></param>
    /// <param name="qtdCapitulos"></param>
    /// <param name="text"></param>
    public Livro(Capitulo? capitulos, string? name, string? abbrev, int qtdCapitulos, string? text )
    {
        Capitulos = capitulos;
        Name = name;
        Abbre = abbrev;
        QtdCapitulos = qtdCapitulos;
        Text = text;
    }

    public Livro()
    {
    }
}