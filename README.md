##About
TabulaRasa is .NET library which provides a fluent API for generating, changing and templating documents in OpenXML format.

##NuGet Install

<img src="http://eskat0n.github.com/images/Nuget-Foxby.png" alt="Nuget Foxby">

NuGet link https://nuget.org/packages/TabulaRasa

##Object Hierarchy

<img src="http://eskat0n.github.com/images/Foxby-Hierarchy.png" alt="Hierarchy">

##Examples

###1 Hello Word!
```csharp
private static void Main()
{
    using (var docxDocument = new DocxDocument(SimpleTemplate.EmptyWordFile))
    {
        var builder = new DocxDocumentBuilder(docxDocument);
 
        builder.Tag(SimpleTemplate.ContentTagName,
                    x => x.Center.Paragraph(z => z.Bold.Text("Hello Word!")));


        File.WriteAllBytes(string.Format(@"D:\Word.docx"), docxDocument.ToArray());
    }
}
```
###2 Table and formatting
```csharp
private static void Main()
{
    string customerName = "Jonh Smith";
    string orderNumber = "4";
    string itemName1 = "Pen";
    string itemSumm1 = "5 000";
    string itemName2 = "Laptop";
    string itemSumm2 = "6 342";
    string summ = "11 342";

    using (var docxDocument = new DocxDocument(SimpleTemplate.EmptyWordFile))
    {
        var builder = new DocxDocumentBuilder(docxDocument);
 
        builder.Tag(SimpleTemplate.ContentTagName,
                    x => x.Center.Paragraph(z => z.Bold.Text(string.Format("Offer №{0}", orderNumber)))
                          .Right.Paragraph(DateTime.Now.ToString("dd MMMM yyyy"))
                          .Left.Paragraph(string.Format("I, {0}, buy:", customerName))
                          .Table(t => t.Column("Item", 70).Column("Price", 30),
                                 r => r.Row(itemName1, itemSumm1)
                                       .Row(itemName2, itemSumm2)
                                       .Row(w => w.Right.Bold.Text("Total:"), 
                                 w => w.Center.Bold.Underlined.Text(summ))));
 

        File.WriteAllBytes(string.Format(@"D:\Word.docx"), docxDocument.ToArray());
    }
}
```
 
