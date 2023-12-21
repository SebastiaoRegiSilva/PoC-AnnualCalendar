using System.Xml.Serialization;

/// <summary>
/// Versículo bíblico.
/// </summary>
[Serializable]
public class Versiculo
{
    [XmlElement("n")]
    public int NumeroDoVersiculo { get; set; }
    
    [XmlElement("text")]
    public string? Text { get; set; }
    
    /// <summary>
    /// Construtor de objetos com parâmetros.
    /// </summary>
    /// <param name="numeroDoVersiculo"></param>
    /// <param name="text"></param>
    public Versiculo(int numeroDoVersiculo, string? text)
    {
        NumeroDoVersiculo = numeroDoVersiculo;
        Text = text;
    }

    public Versiculo()
    {
    }
}