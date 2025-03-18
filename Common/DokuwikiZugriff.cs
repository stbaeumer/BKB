using CookComputing.XmlRpc;
using System.Text;
using Microsoft.Extensions.Configuration;

public struct StructSchema
{
    public string name;
    public StructField[] fields;
}

public struct StructField
{
    public string name;
    public string label;
    public string type;
    public string multi;
    
}

public interface IDokuWikiApi : IXmlRpcProxy
{
    [XmlRpcMethod("struct.getData")]
    XmlRpcStruct GetData(string schemaName, string pageId);
    
    [XmlRpcMethod("struct.getSchema")]
    StructSchema GetSchema(string schemaName);
    
    [XmlRpcMethod("dokuwiki.getVersion")]
    string GetVersion();

    [XmlRpcMethod("wiki.getPage")]
    string GetPage(string page);

    [XmlRpcMethod("wiki.putPage")]
    bool PutPage(string page, string content, XmlRpcStruct options);
    
    
    
}

public class DokuwikiZugriff
{
    public DokuwikiZugriff(IConfiguration configuration)
    {
        Global.Konfig("WikiUrl", configuration, "URL zum dokuwiki xmlrpc.",Global.Datentyp.Url);
        Global.Konfig("WikiJsonUser", configuration, "Benutzer, mit dem auf Json zugegriffen wird.");
        Global.Konfig("WikiJsonUserKennwort", configuration, "Kennwort");
        
        // Proxy erstellen
        Proxy = XmlRpcProxyGen.Create<IDokuWikiApi>();
        ((XmlRpcClientProtocol)Proxy).Url = Global.WikiUrl;

        // Manuelle HTTP-Header setzen
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{Global.WikiJsonUser}:{Global.WikiJsonUserKennwort}"));
        ((XmlRpcClientProtocol)Proxy).Headers.Add("Authorization", "Basic " + credentials);
    }

    public IDokuWikiApi Proxy { get; set; }
    public XmlRpcStruct Options { get; set; }

    public void GetVersion()
    {
        var version = Proxy.GetVersion();
        Console.WriteLine($"DokuWiki Version: {version}");
    }

    public string GetPage(string page)
    {
        var pageContent = Proxy.GetPage("start");
        //Console.WriteLine($"Seiteninhalt: {pageContent}");
        return pageContent;
    }
    
    public void PutPage(string page, string content)
    {
        Proxy.PutPage(page, content, new XmlRpcStruct());
        Console.WriteLine("Seite aktualisiert!");
    }
}