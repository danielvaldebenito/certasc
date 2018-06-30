using System.Linq;
using System.Web;
using System.Diagnostics;
using Newtonsoft.Json;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.IO;
/// <summary>
/// Descripción breve de CreatePDF
/// </summary>
public class CreatePDFD118
{
    public static Document document;
    public static Inspeccion Inspeccion { get; set; }
    public string FileName { get; set; }
    int point = 1;
    int subpoint = 1;
    int page = 1;
    List<BookMark> bookMarkList = new List<BookMark>();
    public static int NormaPrincipal = 21;
    public static int TipoInforme = 6;
    public static string NormaPrincipalNombre = "NCh3344/2:2013";
    public string Rendered { get; set; }
    public CreatePDFD118(Inspeccion inspeccion)
    {
        Inspeccion = inspeccion;
        FileName = "Inspeccion IT " + Inspeccion.IT.Replace('/', '-') + ".pdf";
        document = new Document();
        document.Info.Title = "Inspección";
        document.DefaultPageSetup.TopMargin = "7cm";
        document.DefaultPageSetup.LeftMargin = "2cm";
        document.DefaultPageSetup.RightMargin = "2cm";
        DefineStyles(document);
        DefineCover(document);
        CreateVineta();
        DefineContentSection(document);
        BreveIntroAndAlcance();
        Referencias();
        Antecedentes();
        ImagenCabina();
        TerminosYDefiniciones();
        ResultadosInspeccion();
        ObservacionesNormativasYTecnicas();
        Conclusiones();
        
        DefineTableOfContents(document);
        Rendered = Rendering();
    }
    public static void DefineStyles(Document document)
    {
        // Get the predefined style Normal.
        Style style = document.Styles["Normal"];
        // Because all styles are derived from Normal, the next line changes the
        // font of the whole document. Or, more exactly, it changes the font of
        // all styles and paragraphs that do not redefine the font.
        style.Font.Name = "Arial";
        // Heading1 to Heading9 are predefined styles with an outline level. An outline level
        // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks)
        // in PDF.
        style = document.Styles["Heading1"];
        style.Font.Size = 14;
        style.Font.Bold = true;
        style.Font.Color = Colors.DarkBlue;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.PageBreakBefore = true;
        style.ParagraphFormat.SpaceAfter = "1cm";
        style = document.Styles["Heading2"];
        style.ParagraphFormat.PageBreakBefore = false;
        style.Font.Size = 14;
        style.Font.Bold = true;
        style.Font.Color = Colors.DarkBlue;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.SpaceAfter = "1cm";
        style.ParagraphFormat.SpaceBefore = "1cm";

        style = document.Styles["Heading3"];
        style.Font.Size = 10;
        style.Font.Bold = true;
        style.Font.Italic = true;
        style.ParagraphFormat.SpaceBefore = 6;
        style.ParagraphFormat.SpaceAfter = 3;
        style = document.Styles[StyleNames.Header];
        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);
        style = document.Styles[StyleNames.Footer];
        style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);
        // Create a new style called TextBox based on style Normal
        style = document.Styles.AddStyle("TextBox", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        style.ParagraphFormat.Borders.Width = 2.5;
        style.ParagraphFormat.Borders.Distance = "3pt";
        style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

        // Parrafo Normal
        style = document.Styles.AddStyle("Parrafo", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        style.ParagraphFormat.Font.Size = 11;
        style.ParagraphFormat.SpaceBefore = "0.5cm";
        style.ParagraphFormat.SpaceAfter = "0.5cm";

        // Caract
        style = document.Styles.AddStyle("Caract", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        style.ParagraphFormat.Font.Size = 10;
        style.ParagraphFormat.SpaceBefore = "0.2cm";
        style.ParagraphFormat.SpaceAfter = "0.2cm";

        // Pie de fotos
        style = document.Styles.AddStyle("Pie", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.Font.Size = 9;
        style.ParagraphFormat.SpaceBefore = "0.1cm";
        style.ParagraphFormat.SpaceAfter = "0.1cm";
        style.ParagraphFormat.Font.Color = Colors.Blue;
        // Create a new style called TOC based on style Normal
        style = document.Styles.AddStyle("TOC", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 10;
        style.ParagraphFormat.SpaceBefore = "0.3cm";
        style.ParagraphFormat.SpaceAfter = "0.3cm";
        style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
        style.ParagraphFormat.Font.Color = Colors.Black;
    }
    public void DefineCover(Document document)
    {
        Section section = document.AddSection();
        section.PageSetup.TopMargin = 30;
        Paragraph paragraph = section.AddParagraph();

        //paragraph.Format.SpaceAfter = "1cm";
        string pathImage = HttpContext.Current.Server.MapPath("~/css/images/");
        var parr = section.AddParagraph();
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.SpaceBefore = "5cm";
        Image image = section.LastParagraph.AddImage(pathImage + "/logo.png");
        image.Width = "6cm";
        paragraph = section.AddParagraph(string.Format("INFORME DE AUDITORÍA TÉCNICA E INSPECCIÓN DEL {0}", Inspeccion.Aparato.Nombre.ToUpper()));
        paragraph.Format.Font.Size = 16;
        paragraph.Format.Alignment = ParagraphAlignment.Center;
        paragraph.Format.Font.Color = Colors.Black;
        paragraph.Format.Font.Bold = true;
        paragraph.Format.SpaceBefore = "2cm";

        paragraph = section.AddParagraph(string.Format("IT N° {0}", Inspeccion.IT));
        paragraph.Format.Font.Size = 12;
        paragraph.Format.Font.Bold = false;
        paragraph.Format.SpaceAfter = "1cm";
        paragraph.Format.Alignment = ParagraphAlignment.Center;

        paragraph = section.AddParagraph(string.Format("INFORME FASE {0} {1} N° {2}", ToRoman(Inspeccion.Fase), Inspeccion.Aparato.Nombre, Inspeccion.Numero));
        paragraph.Format.Font.Size = 12;
        paragraph.Format.Font.Bold = false;
        paragraph.Format.SpaceAfter = "0.5cm";
        paragraph.Format.Alignment = ParagraphAlignment.Center;

        paragraph = section.AddParagraph(string.Format("Edificio {0}", Inspeccion.NombreEdificio));
        paragraph.Format.Font.Size = 12;
        paragraph.Format.Font.Bold = false;
        paragraph.Format.SpaceAfter = "0.5cm";
        paragraph.Format.Alignment = ParagraphAlignment.Center;

        paragraph = section.AddParagraph(string.Format("Fecha de Inspección {0}", Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy")));
        paragraph.Format.Font.Size = 11;
        paragraph.Format.Font.Bold = false;
        paragraph.Format.SpaceAfter = "7cm";
        paragraph.Format.Alignment = ParagraphAlignment.Center;
    }
    public void CreateVineta()
    {
        //document.LastSection.AddParagraph("Cell Merge", "Heading2");
        Table table = document.LastSection.AddTable();

        table.Borders.Visible = true;
        table.Borders.Width = 1;
        table.Borders.Color = Colors.Gray;
        table.TopPadding = 5;
        table.BottomPadding = 5;
        Column column = table.AddColumn();
        column.Format.Alignment = ParagraphAlignment.Left;
        column.Width = 120;
        column = table.AddColumn();
        column.Width = 120;
        column = table.AddColumn();
        column.Width = 125;
        column = table.AddColumn();
        column.Width = 125;

        table.Rows.Height = 18;
        Row row = table.AddRow();
        row.Cells[0].AddParagraph("SECCIÓN AUDITORÍA E INSPECCIÓN PARA CERTIFICACIÓN – DEPTO. DE INGENIERÍA");
        row.Cells[0].MergeRight = 1;
        row.Cells[0].Shading.Color = Colors.LightGray;
        row.Cells[2].AddParagraph(string.Format("REF. IT: {0}", Inspeccion.IT));

        row.Cells[3].AddParagraph("EJEMPLAR N° 1");
        row.Format.Font.Bold = true;

        row = table.AddRow();
        row.VerticalAlignment = VerticalAlignment.Center;
        row.Format.Font.Bold = true;
        row.Cells[0].AddParagraph("ELABORADO POR");
        row.Cells[1].AddParagraph("REVISADO POR");
        row.Cells[2].AddParagraph("APROBADO POR");
        row.Cells[3].AddParagraph("DESTINATARIO");


        row = table.AddRow();
        row.Cells[0].AddParagraph(string.Format("CARGO: {0} \n {1} {2}", Inspeccion.Usuario.Cargo, Inspeccion.Usuario.Nombre, Inspeccion.Usuario.Apellido));
        row.Cells[1].AddParagraph("Unidad Inspección de Especialidades y Transporte Vertical");
        row.Cells[2].AddParagraph(string.Format("{0} - {1}", Inspeccion.Aprobador == null ? string.Empty : Inspeccion.Usuario1.Cargo, Inspeccion.Aprobador == null ? string.Empty : Inspeccion.Usuario1.Nombre + " " + Inspeccion.Usuario1.Apellido));
        row.Cells[3].AddParagraph(Inspeccion.Destinatario ?? string.Empty);

        row = table.AddRow();
        row.Height = 12;
        row.Cells[0].AddParagraph("FECHA");
        row.Cells[1].AddParagraph("FECHA");
        row.Cells[2].AddParagraph("FECHA");
        row.Cells[3].AddParagraph("FECHA");


        row = table.AddRow();
        var fechaElab = string.Empty;
        if (Inspeccion.Cumplimiento.Any())
        {
            fechaElab = ((DateTime)Inspeccion
                            .Cumplimiento
                            .Min(m => m.Fecha))
                            .ToString("dd-MM-yyyy");
        }
         
        row.Cells[0].AddParagraph(fechaElab);
        row.Cells[1].AddParagraph(Inspeccion.FechaRevision.HasValue ? Inspeccion.FechaRevision.Value.ToString("dd-MM-yyyy") : string.Empty); // Guardar fecha aprobación
        row.Cells[2].AddParagraph(Inspeccion.FechaAprobacion.HasValue ? Inspeccion.FechaAprobacion.Value.ToString("dd-MM-yyyy") : string.Empty);// Guardar fecha aprobación
        row.Cells[3].AddParagraph(Inspeccion.FechaEntrega.HasValue ? Inspeccion.FechaEntrega.Value.ToString("dd-MM-yyyy") : string.Empty); // Guardar fecha entrega


    }
    public void DefineContentSection(Document document)
    {
        
        HeaderFooter header = document.LastSection.Headers.Primary;
        header.Format.Alignment = ParagraphAlignment.Center;
        header.Format.SpaceAfter = "1cm";

        Table tableHeader = header.AddTable();
        tableHeader.Borders.Visible = true;
        tableHeader.Borders.Color = Colors.LightGray;
        tableHeader.Format.Font.Color = Colors.DarkSlateGray;
        Column column = tableHeader.AddColumn();
        column.Width = 75;
        column = tableHeader.AddColumn();
        column.Width = 95;
        column = tableHeader.AddColumn();
        column.Width = 150;
        column = tableHeader.AddColumn();
        column.Width = 95;
        column = tableHeader.AddColumn();
        column.Width = 75;
        Row row = tableHeader.AddRow();
        row.HeadingFormat = true;
        row.Height = 50;
        string pathImage = HttpContext.Current.Server.MapPath("~/css/images/");
        Image image = row.Cells[0].AddImage(pathImage + "/logo.png");
        image.Width = 60;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.Cells[1].AddParagraph("INSPECCIÓN NORMA NCh3344/2:2013\n Lista de comprobación para la seguridad de escaleras mecánicas y de las rampas móviles existentes.");
        row.Cells[1].MergeRight = 2;
        row.Cells[4].Format.Alignment = ParagraphAlignment.Center;
        var p = row.Cells[4].AddParagraph();
        p.Format.Alignment = ParagraphAlignment.Center;
        image = p.AddImage(pathImage + "/logo_sgs.png");
        image.Width = 40;
        row = tableHeader.AddRow();
        Paragraph parrafo = row.Cells[0].AddParagraph("SECCIÓN GESTIÓN E INSPECCIÓN DE PROYECTOS - SGP");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 10;
        row.Cells[0].MergeRight = 1;
        row.Cells[0].MergeDown = 5;
        row.Cells[2].AddParagraph(string.Format("INSPECCIÓN FASE {0} \n Check List \n Inspección de escaleras mecánicas o rampas móviles", ToRoman(Inspeccion.Fase)));
        row.Cells[2].MergeDown = 6;

        parrafo = row.Cells[3].AddParagraph("VERSIÓN");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph(Inspeccion.FechaAprobacion.HasValue ? "1.0" : "Preliminar");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row = tableHeader.AddRow();
        parrafo = row.Cells[3].AddParagraph("FECHA");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph("05-07-2016");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        
        row = tableHeader.AddRow();
        parrafo = row.Cells[3].AddParagraph("Revisado por ");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph("H.B.V.");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row = tableHeader.AddRow();
        parrafo = row.Cells[3].AddParagraph("Aprobado por ");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph("M.J.M.");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row = tableHeader.AddRow();
        parrafo = row.Cells[3].AddParagraph("Fecha Aprobación");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph("08-07-2016");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row = tableHeader.AddRow();
        parrafo = row.Cells[3].AddParagraph("Código");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph("DI - 118");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row = tableHeader.AddRow();
        parrafo = row.Cells[0].AddParagraph("Elaborado por");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[1].AddParagraph("D. Ingeniería Certel");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[3].AddParagraph("Página");
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        parrafo = row.Cells[4].AddParagraph();
        parrafo.AddPageField();
        parrafo.Format.Font.Bold = true;
        parrafo.Format.Font.Size = 8;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;


        // Footer
        HeaderFooter footer = document.LastSection.Footers.Primary;
        Table tableFooter = footer.AddTable();
        tableFooter.Borders.Visible = false;
        Column col = tableFooter.AddColumn();
        col.Width = 80;
        col = tableFooter.AddColumn();
        col.Width = 300;
        col = tableFooter.AddColumn();
        col.Width = 80;
        col = tableFooter.AddColumn();
        col.Width = 30;
        Row row1 = tableFooter.AddRow();
        var parr = row1.Cells[0].AddParagraph("www.certel.cl");
        row1.VerticalAlignment = VerticalAlignment.Center;
        parr.Format.Font.Size = 8;
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Color = Colors.Blue;
        parr = row1.Cells[1].AddParagraph("Carlos Condell N° 198, Buin, Santiago / Tel.: (+56) 223005921 - 223028182");
        parr.Format.Font.Size = 6;
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr = row1.Cells[2].AddParagraph("contacto@certel.cl");
        parr.Format.Font.Color = Colors.Blue;
        parr.Format.Font.Size = 8;
        parr.Format.Alignment = ParagraphAlignment.Center;
        // Create a paragraph with centered page number. See definition of style "Footer".
        Paragraph paragraph = new Paragraph();
        paragraph.AddTab();
        paragraph.AddPageField();
        paragraph.Format.Font.Size = 12;
        row1.Cells[3].Add(paragraph);
        row1.Cells[3].VerticalAlignment = VerticalAlignment.Center;
        parr.Format.Alignment = ParagraphAlignment.Center;
        // Add clone of paragraph to footer for odd pages. Cloning is necessary because an object must
        // not belong to more than one other object. If you forget cloning an exception is thrown.
        //section.Footers.EvenPage.Add(paragraph.Clone());
        page++;
    }

    public string ToRoman(int number)
    {
        switch (number)
        {
            case 1: return "I";
            case 2: return "II";
            case 3: return "III";
            default: return number.ToString();
        }
    }
    public void BreveIntroAndAlcance()
    {
        Section section = document.AddSection();
        //section.PageSetup.TopMargin = 200;
        Paragraph title = section.AddParagraph(string.Format("INSPECCIÓN DE {0} N° {1} EDIFICIO {2}", Inspeccion.Aparato.Nombre.ToUpper(), Inspeccion.Numero, Inspeccion.NombreEdificio)); // Nombre edificio reemplazar
        title.Style = "Heading2";

        Paragraph parrafo = section.AddParagraph(string.Format("El presente informe se refiere a los resultados de la Inspección de la Auditoría Técnica denominada Fase {0}, realizada a la instalación de la {1}, ubicada en {2}.", ToRoman(Inspeccion.Fase), Inspeccion.Aparato.Nombre, Inspeccion.Ubicacion));
        parrafo.Style = "Parrafo";
        title = section.AddParagraph(string.Format("{0}. ALCANCE", point));
        title.Style = "Heading2";
        title.AddBookmark("alcance");
        bookMarkList.Add(new BookMark
        {
            Text = string.Format("{0}. ALCANCE", point),
            Mark = "alcance",
            IsSub = false
        });
        subpoint = 1;
        Paragraph parrafo1 = section.AddParagraph(string.Format("{0} \tEl presente documento tiene por objeto informar las observaciones necesarias a resolver para lograr la certificación de la {1} del edificio de la referencia, y establecer los requisitos de seguridad que debe cumplir por norma cada {1}, con el fin de proteger a las personas y objetos contra riesgos de accidentes durante su funcionamiento y también mientras se realizan los trabajos de mantenimiento e inspección.", point + "." + subpoint + ".", Inspeccion.Aparato.Nombre));
        parrafo1.Style = "Parrafo";
        subpoint++;
        parrafo1 = section.AddParagraph(string.Format("{0} \tEl objetivo es verificar el estado actual de la {1}, respecto de la norma NCh3344/2:2013 y ver si cumple con los requisitos de construcción e instalación de la misma; comprobar que se mantiene en condiciones de funcionamiento seguro, tanto para los pasajeros o usuarios, como para el personal que realiza el servicio de mantenimiento, para así comprobar si ésta se puede certificar", point + "." + subpoint + ".", Inspeccion.Aparato.Nombre));
        subpoint++;
        parrafo1.Style = "Parrafo";
        parrafo1 = section.AddParagraph(string.Format("{0} \t“Por no existir Norma Internacional, en la elaboración de esta norma se ha tomado en consideración la norma UNE-EN 115-2:2011 Seguridad de escaleras mecánicas y andenes móviles - Parte 2: Reglas para la seguridad de las escaleras mecánicas y de los andenes móviles existentes”. Normalizada por la INN (Instituto Nacional de Normalización) NCh3344/2:2013 con el mismo nombre.", point + "." + subpoint + "."));
        subpoint++;
        parrafo1.Style = "Parrafo";
        parrafo1 = section.AddParagraph(string.Format("{0} \tCertel, en su totalidad y como organismo de inspección, tiene completa independencia e imparcialidad respecto a sus clientes, funcionando como Organismo de Inspección bajo los criterios de la norma chilena NCh ISO 17.020.", point + "." + subpoint + "."));
        parrafo1.Style = "Parrafo";

    }
    public void Referencias()
    {
        Section section = document.AddSection();
        point++;
        subpoint = 1;
        Paragraph title = section.AddParagraph(string.Format("{0}. REFERENCIAS", point));
        title.Style = "Heading1";
        title.AddBookmark("referencias");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}. REFERENCIAS", point), Mark = "referencias", IsSub = false });
        Paragraph texto = section.AddParagraph("En la evaluación se utilizó como referencia las siguientes normas:");
        texto.Style = "Parrafo";
        texto = section.AddParagraph(string.Format("{0}.{1}	Instituto nacional de normalización, INN: NCh3344/1:2013, \"Seguridad de escaleras mecánicas y rampas móviles - Parte 1: Contrucción e instalación\".", point, subpoint));
        texto.Style = "Parrafo";
        subpoint++;
        texto = section.AddParagraph(string.Format("{0}.{1}	Instituto nacional de normalización, INN: NCh3344/2:2013, \"Seguridad de escaleras mecánicas y rampas móviles - Parte 1: Contrucción e instalación\".", point, subpoint));
        texto.Style = "Parrafo";
        subpoint++;
        texto = section.AddParagraph(string.Format("{0}.{1}	Superintendencia de electricidad y combustibles, NCh Elec. 4/2003, \"Instalaciones de consumo en baja tensión\"", point, subpoint));
        texto.Style = "Parrafo";
        subpoint++;
        
        texto = section.AddParagraph(string.Format("{0}.{1}	NCh ISO 17020:2012 - Evaluación de la conformidad - Requisitos para el funcionamiento de los diversos tipos de organismo que realizan inspección.", point, subpoint));
        texto.Style = "Parrafo";
        subpoint++;
        texto = section.AddParagraph(string.Format("{0}.{1}	D. S. N° 47 Ordenanza General de Urbanismo y Construcciones (Actualizada al 21 de Marzo 2016 - incorpora modificaciones D. S. N° 50 D.O. 04-03-2016 - D. S. N° 37 - D.O. 21-03-2016).", point, subpoint));
        texto.Style = "Parrafo";

    }
    public void Antecedentes()
    {
        Section section = document.AddSection();
        point++;
        Paragraph title = section.AddParagraph(string.Format("{0}. ANTECEDENTES", point));
        title.Style = "Heading1";
        title.AddBookmark("antecedentes");
        bookMarkList.Add(new BookMark
        {
            Text = string.Format("{0}. ANTECEDENTES", point),
            Mark = "antecedentes",
            IsSub = false
        });
        title.Format.SpaceAfter = 2;
        Paragraph parrafo = section.AddParagraph(string.Format("En esta inspección se verifica el cumplimiento de la norma NCh3344/2:2013, asociada a las instalaciones y el funcionamiento del {0}.", Inspeccion.Aparato.Nombre));
        parrafo.Style = "Parrafo";

        Paragraph tableTitle = section.AddParagraph("TABLA N°1");
        tableTitle.Style = "Heading2";
        subpoint = 1;
        Table table1 = section.AddTable();
        table1.Borders.Visible = true;
        table1.Borders.Color = Colors.LightGray;
        Column column = table1.AddColumn();
        column.Width = 200;
        column = table1.AddColumn();
        column.Width = 90;
        column = table1.AddColumn();
        column.Width = 200;
        Row row = table1.AddRow();
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;
        row.Cells[0].MergeRight = 2;
        Paragraph parrafo1 = row.Cells[0].AddParagraph(string.Format("EQUIPOS Y HERRAMIENTAS UTILIZADOS"));
        row.Cells[0].Shading.Color = Colors.LightGray;
        parrafo1.AddBookmark("equipos");
        bookMarkList.Add(new BookMark
        {
            Text = "TABLA N° 1",
            Mark = "equipos",
            IsSub = true
        });
        row = table1.AddRow();
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;
        row.Cells[0].AddParagraph("TIPO");
        row.Cells[1].AddParagraph("N° IDENT");
        row.Cells[2].AddParagraph("IDENTIFICACIÓN");

        using (var db = new CertelEntities())
        {
            var equipos = db.EquipoUtilizado
                            .Where(w => w.Usuario == Inspeccion.Ingeniero)
                            .ToList();
            foreach (var eq in equipos)
            {
                Row rowe = table1.AddRow();
                rowe.VerticalAlignment = VerticalAlignment.Center;
                rowe.TopPadding = 2;
                rowe.BottomPadding = 2;
                rowe.Cells[0].AddParagraph(eq.Tipo);
                rowe.Cells[1].AddParagraph(eq.Ident);
                rowe.Cells[2].AddParagraph(eq.Identificacion);
            }
        }

        // Tabla 2

        var insp = Inspeccion.Fase == 1 ? Inspeccion : Inspeccion.Inspeccion2;
        section.AddPageBreak();
        tableTitle = section.AddParagraph("TABLA N°2");
        tableTitle.Style = "Heading2";
        Table table2 = section.AddTable();
        table2.Borders.Visible = true;
        table2.KeepTogether = true;
        table2.Borders.Color = Colors.LightGray;
        Column column2 = table2.AddColumn();
        column2.Width = 245;
        column2 = table2.AddColumn();
        column2.Width = 245;

        Row row2 = table2.AddRow();
        row2.Format.Font.Bold = true;
        row2.Format.Alignment = ParagraphAlignment.Center;
        row2.VerticalAlignment = VerticalAlignment.Center;
        row2.TopPadding = 5;
        row2.BottomPadding = 5;
        row2.Cells[0].MergeRight = 1;

        parrafo1 = row2.Cells[0].AddParagraph("CARACTERÍSTICAS GENERALES");
        row2.Cells[0].Shading.Color = Colors.LightGray;
        parrafo1.AddBookmark("caracteristicas generales");
        bookMarkList.Add(new BookMark
        {
            Text = "TABLA N° 2",
            Mark = "caracteristicas generales",
            IsSub = true
        });
        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Nombre del Proyecto");
        row2.Cells[1].AddParagraph(insp.NombreProyecto ?? string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Ubicación");
        row2.Cells[1].AddParagraph(insp.Ubicacion);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Destino del Proyecto");
        row2.Cells[1].AddParagraph(insp.DestinoProyectoID == null ? string.Empty : insp.DestinoProyecto.Descripcion);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Altura en pisos");
        row2.Cells[1].AddParagraph(insp.AlturaPisos.ToString());
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Permiso Edificación");
        row2.Cells[1].AddParagraph(insp.PermisoEdificacion ?? "Sin información");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Recepción Municipal");
        row2.Cells[1].AddParagraph(insp.RecepcionMunicipal ?? "Sin información");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Número único del elevador");
        row2.Cells[1].AddParagraph(insp.Numero ?? string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        
        // Especificos Tabla 2
        var especificosT2 = insp.ValoresEspecificos.Where(w => w.Especificos.NroTable == 2).OrderBy(o => o.EspecificoID);
        foreach (var e in especificosT2)
        {
            row2 = table2.AddRow();
            row2.Cells[0].AddParagraph(e.Especificos.Nombre);
            row2.Cells[1].AddParagraph(e.Valor);
            row2.TopPadding = 2;
            row2.BottomPadding = 2;
        }

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Fecha de Emisión del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph(insp.FechaEmisionCertificado.HasValue ? insp.FechaEmisionCertificado.Value.ToString("dd-MM-yyyy") : string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Fecha de Vencimiento del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph(insp.FechaVencimientoCertificado.HasValue ? insp.FechaVencimientoCertificado.Value.ToString("dd-MM-yyyy") : string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        // Especificos
        section.AddPageBreak();
        section.AddParagraph();
        tableTitle = section.AddParagraph("TABLA N°3");
        tableTitle.Style = "Heading2";
        Table table3 = section.AddTable();
        table3.KeepTogether = true;
        table3.Borders.Visible = true;
        table3.Borders.Color = Colors.LightGray;
        Column column3 = table3.AddColumn();
        column3.Width = 245;
        column3 = table3.AddColumn();
        column3.Width = 245;

        Row row3 = table3.AddRow();
        row3.Format.Font.Bold = true;
        row3.Format.Alignment = ParagraphAlignment.Center;
        row3.VerticalAlignment = VerticalAlignment.Center;
        row3.TopPadding = 5;
        row3.BottomPadding = 5;
        row3.Cells[0].MergeRight = 1;

        parrafo1 = row3.Cells[0].AddParagraph("CARACTERÍSTICAS PARTICULARES");
        parrafo1.AddBookmark("caracteristicas particulares");
        bookMarkList.Add(new BookMark
        {
            Text = "TABLA N° 3",
            Mark = "caracteristicas particulares",
            IsSub = true
        });
        row3.Shading.Color = Colors.LightGray;
        row3.Borders.Color = Colors.White;
        row3 = table3.AddRow();
        row3.Cells[0].AddParagraph("CARACTERÍSTICAS DEL EQUIPO");
        row3.Cells[1].AddParagraph(insp.Aparato.Nombre + " N° " + insp.Numero);
        row3.Shading.Color = Colors.LightGray;
        row3.Borders.Color = Colors.White;
        row3.Format.Font.Bold = true;
        row3.TopPadding = 5;
        row3.BottomPadding = 5;
        row3.Format.Alignment = ParagraphAlignment.Center;
        row3.VerticalAlignment = VerticalAlignment.Center;


        var especificos = insp.ValoresEspecificos.Where(w => w.Especificos.NroTable == 3).OrderBy(o => o.EspecificoID);
        foreach (var esp in especificos)
        {
            row3 = table3.AddRow();
            row3.Cells[0].AddParagraph(esp.Especificos.Nombre);
            row3.Cells[1].AddParagraph(esp.Valor);
            row3.TopPadding = 2;
            row3.BottomPadding = 2;
        }
    }
    public void ImagenCabina()
    {
        Section section = document.AddSection();
        subpoint = 1;
        Paragraph title = section.AddParagraph(string.Format("{0}.{1} ÁREA TIPO DE LA AUDITORÍA DEL {2}", point, subpoint, Inspeccion.Aparato.Nombre.ToUpper()));
        title.Style = "Heading1";
        title.Format.SpaceAfter = "2cm";
        title.AddBookmark("imagen");
        bookMarkList.Add(new BookMark
        {
            Text = string.Format("{0}.{1} ÁREA TIPO DE LA AUDITORÍA DEL {2}", point, subpoint, Inspeccion.Aparato.Nombre.ToUpper()),
            Mark = "imagen",
            IsSub = true
        });
        string pathImage = HttpContext.Current.Server.MapPath("~/css/images/");
        Image image = section.LastParagraph.AddImage(pathImage + "/cabina112.png");

    }
    public void TerminosYDefiniciones()
    {
         subpoint++;
        Section section = document.AddSection();
        Paragraph title = section.AddParagraph(string.Format("{0}.{1}. ALGUNOS TÉRMINOS Y DEFINICIONES", point, subpoint));
        title.Style = "Heading1";
        title.AddBookmark("terminos");
       
        bookMarkList.Add(new BookMark
        {
            Text = string.Format("{0}.{1}. ALGUNOS TÉRMINOS Y DEFINICIONES", point, subpoint),
            Mark = "terminos",
            IsSub = true
        });
        using (var db = new CertelEntities())
        {
            var terminos = db.TerminosYDefiniciones
                            .Where(w => w.NormaID == NormaPrincipal)
                            .ToList();
            foreach (var t in terminos)
            {
                Paragraph termino = section.AddParagraph(t.Termino);
                termino.Format.Font.Size = 11;
                termino.Format.Font.Bold = true;
                termino.Format.SpaceAfter = 2;
                Paragraph definicion = section.AddParagraph(t.Definicion.TrimEnd());
                definicion.Style = "Parrafo";
                definicion.Format.LeftIndent = "1cm";
                definicion.Format.SpaceBefore = "0.1cm";
                definicion.Format.Alignment = ParagraphAlignment.Justify;
            }
        }
    }
    public void ResultadosInspeccion()
    {
        Section section = document.AddSection();
        point++;
        subpoint = 1;
        var subsubpoint = 1;
        Paragraph title = section.AddParagraph(string.Format("{0}. RESULTADOS DE LA INSPECCIÓN DEL {1}", point, Inspeccion.Aparato.Nombre.ToUpper()));
        title.Style = "Heading1";
        title.AddBookmark("resultados");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}. RESULTADOS DE LA INSPECCIÓN DEL {1}", point, Inspeccion.Aparato.Nombre.ToUpper()), Mark = "resultados", IsSub = false });
        Paragraph texto = section.AddParagraph(string.Format("A continuación se verifican las áreas de inspección y se detallan las no conformidades encontradas tras la Fase {0} del proceso de certificación en el equipo referente a la norma NCh3344/2:2013, respecto a la lista de verificación técnica de la misma, las que deben ser tratadas por seguridad y para poder optar a la certificación de la {1}.", ToRoman(Inspeccion.Fase), Inspeccion.Aparato.Nombre));
        texto.Style = "Parrafo";
        texto = section.AddParagraph("GLOSARIO");
        texto.Style = "Parrafo";
        texto.Format.Font.Bold = true;
        using (var db = new CertelEntities())
        {
            var glosario = db.Evaluacion
                            .Where(w => w.Fase == 1);
            foreach (var g in glosario)
            {
                texto = section.AddParagraph(string.Format("{0}: {1}", g.Glosa, g.Descripcion));
                texto.Style = "Parrafo";
            }

            var n = Inspeccion
                            .InspeccionNorma
                            .Where(w => !w.Norma.NormasAsociadas1.Any())
                            .Where(w => w.Norma.TipoInformeID == TipoInforme)
                            .Select(s => s.Norma)
                            .FirstOrDefault();
            if (n == null)
                return;

            var titulos = n.Titulo;
            foreach (var t in titulos)
            {

                subsubpoint = 1;
                title = section.AddParagraph(string.Format("{0}.{1}. {2}", point, subpoint, t.Texto.ToUpper()));
                title.Style = "Heading2";
                title.AddBookmark(string.Format("titulo{0}", subpoint));
                bookMarkList.Add(new BookMark
                {
                    Text = string.Format("{0}.{1}. {2}", point, subpoint, t.Texto.ToUpper()),
                    Mark = string.Format("titulo{0}", subpoint),
                    IsSub = true
                });
                Table table = section.AddTable();
                table.Borders.Visible = true;
                table.Borders.Color = Colors.LightGray;
                table.KeepTogether = false;
                Column column = table.AddColumn();
                column.Width = 35;
                column = table.AddColumn();
                column.Width = 75;
                column = table.AddColumn();
                column.Width = 230;
                column = table.AddColumn();
                column.Width = 30;
                column = table.AddColumn();
                column.Width = 120;
                Row row = table.AddRow();
                row.TopPadding = 5;
                row.BottomPadding = 5;
                row.Format.Font.Bold = true;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = VerticalAlignment.Center;
                row.Cells[0].MergeRight = 1;
                row.Cells[0].AddParagraph(string.Format("{0}", n.Nombre));

                row.Cells[2].MergeDown = 1;
                row.Cells[2].AddParagraph(string.Format("{0}", n.TituloRegulacion));

                row.Cells[3].MergeRight = 1;
                row.Cells[3].AddParagraph("CUMPLIMIENTO");

                row = table.AddRow();
                row.Format.Font.Bold = true;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = VerticalAlignment.Center;
                row.TopPadding = 5;
                row.BottomPadding = 5;
                row.Cells[0].AddParagraph("IDENT");
                row.Cells[1].AddParagraph("REQUISITO");
                row.Cells[3].AddParagraph("OK N/A N/C");
                row.Cells[4].AddParagraph("OBSERVACIONES");
                var requisitos = t.Requisito.Where(w => w.Habilitado == true);
                foreach (var r in requisitos)
                {
                    var cars = r.Caracteristica.Where(w => w.Habilitado == true);
                    var carsCount = cars.Count();
                    if (carsCount == 0)
                        continue;
                    foreach (var c in cars)
                    {
                        var cRow = table.AddRow();
                        cRow.Format.Alignment = ParagraphAlignment.Center;
                        cRow.VerticalAlignment = VerticalAlignment.Center;
                        cRow.TopPadding = 0;
                        cRow.BottomPadding = 0;
                        var parr1 = cRow.Cells[2].AddParagraph(c.Descripcion);
                        parr1.Style = "Caract";
                        cRow.Cells[0].AddParagraph(string.Format("{0}.{1}.{2}", point, subpoint, subsubpoint));
                        cRow.Cells[1].AddParagraph(string.Format("{0}", r.Descripcion));
                        cRow.Cells[0].MergeDown = carsCount - 1;
                        cRow.Cells[1].MergeDown = carsCount - 1;


                        var cumplimiento = c.Cumplimiento
                                            .Where(w => Inspeccion.Fase == 1
                                                    ? w.InspeccionID == Inspeccion.ID
                                                    : w.InspeccionID == Inspeccion.InspeccionFase1)
                                            .FirstOrDefault();
                        if (cumplimiento == null)
                            continue;
                        parr1 = cRow.Cells[3].AddParagraph(cumplimiento == null ? string.Empty : cumplimiento.Evaluacion.Glosa);
                        parr1.Style = "Parrafo";
                        parr1.Format.Alignment = ParagraphAlignment.Center;
                        parr1 = cRow.Cells[4].AddParagraph(cumplimiento.Observacion ?? string.Empty);
                        parr1.Style = "Carac";
                        parr1.Format.Alignment = ParagraphAlignment.Justify;
                        if (cumplimiento.EvaluacionID == 3)
                        {
                            parr1.Format.Font.Color = Colors.Blue;
                            if (Inspeccion.Fase > 1)
                            {
                                var corregido = c.Cumplimiento
                                                    .Where(w => w.InspeccionID == Inspeccion.ID)
                                                    .FirstOrDefault();
                                if (corregido != null)
                                {
                                    parr1 = cRow.Cells[4].AddParagraph(string.Format("{0} en Fase {1}", corregido.Evaluacion.Descripcion, ToRoman(Inspeccion.Fase)));
                                    parr1.Style = "Carac";
                                    parr1.Format.Alignment = ParagraphAlignment.Justify;
                                    parr1.Format.Font.Color = Colors.Blue;
                                    parr1.Format.Shading.Color = Colors.Yellow;
                                }

                            }

                        }


                    }
                    subsubpoint++;
                }
                subpoint++;
            }
            var normasAsociadas = n.NormasAsociadas;

            foreach (var nor in normasAsociadas)
            {
                var na = db.Norma.Find(nor.NormaSecundariaID);
                var titleShowed = false;
                var titles = na.Titulo.ToList();
                foreach (var t in titles)
                {

                    subsubpoint = 1;
                    title = section.AddParagraph(string.Format("{0}.{1}. {2}", point, subpoint, t.Texto.ToUpper()));
                    title.Style = "Heading1";
                    title.AddBookmark(string.Format("titulo{0}", subpoint));
                    bookMarkList.Add(new BookMark
                    {
                        Text = string.Format("{0}.{1}. {2}", point, subpoint, t.Texto.ToUpper()),
                        Mark = string.Format("titulo{0}", subpoint),
                        IsSub = true
                    });
                    if (!titleShowed && na.ParrafoIntroductorio != null)
                    {
                        var parr = section.AddParagraph(na.ParrafoIntroductorio);
                        parr.Style = "Parrafo";
                        titleShowed = true;
                    }
                    Table table = section.AddTable();
                    table.Borders.Visible = true;
                    table.Borders.Color = Colors.LightGray;
                    table.Format.KeepTogether = false;
                    Column column = table.AddColumn();
                    column.Width = 35;
                    column = table.AddColumn();
                    column.Width = 75;
                    column = table.AddColumn();
                    column.Width = 230;
                    column = table.AddColumn();
                    column.Width = 30;
                    column = table.AddColumn();
                    column.Width = 120;
                    Row row = table.AddRow();
                    row.Format.Font.Bold = true;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = VerticalAlignment.Center;
                    row.TopPadding = 5;
                    row.BottomPadding = 5;
                    row.Cells[0].MergeRight = 1;
                    row.Cells[0].AddParagraph(string.Format("{0}", na.Nombre));

                    row.Cells[2].MergeDown = 1;
                    row.Cells[2].AddParagraph(string.Format("{0}", na.TituloRegulacion));

                    row.Cells[3].MergeRight = 1;
                    row.Cells[3].AddParagraph("CUMPLIMIENTO");

                    row = table.AddRow();
                    row.Format.Font.Bold = true;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = VerticalAlignment.Center;
                    row.Cells[0].AddParagraph("IDENT");
                    row.Cells[1].AddParagraph("REQUISITO");
                    row.Cells[3].AddParagraph("OK N/A N/C");
                    row.Cells[4].AddParagraph("OBSERVACIONES");
                    var reqs = t.Requisito.Where(w => w.Habilitado == true).ToList();
                    foreach (var r in reqs)
                    {
                        var cars = r.Caracteristica.Where(w => w.Habilitado == true);
                        var carsCount = cars.Count();
                        if (carsCount == 0)
                            continue;
                        foreach (var c in cars)
                        {
                            var cRow = table.AddRow();
                            cRow.Format.Alignment = ParagraphAlignment.Center;
                            cRow.VerticalAlignment = VerticalAlignment.Center;
                            cRow.TopPadding = 0;
                            cRow.BottomPadding = 0;
                            var parr1 = cRow.Cells[2].AddParagraph(c.Descripcion);
                            parr1.Style = "Caract";
                            cRow.Cells[0].AddParagraph(string.Format("{0}.{1}.{2}", point, subpoint, subsubpoint));
                            cRow.Cells[1].AddParagraph(string.Format("{0}", r.Descripcion));


                            cRow.Cells[0].MergeDown = carsCount - 1;
                            cRow.Cells[1].MergeDown = carsCount - 1;

                            var cumplimiento = c.Cumplimiento
                                            .Where(w => Inspeccion.Fase == 1 ? w.InspeccionID == Inspeccion.ID : w.InspeccionID == Inspeccion.InspeccionFase1).FirstOrDefault();
                            if (cumplimiento == null)
                                continue;
                            parr1 = cRow.Cells[3].AddParagraph(cumplimiento == null ? string.Empty : cumplimiento.Evaluacion.Glosa);
                            parr1.Style = "Parrafo";
                            parr1.Format.Alignment = ParagraphAlignment.Center;
                            parr1 = cRow.Cells[4].AddParagraph(cumplimiento.Observacion ?? string.Empty);
                            parr1.Style = "Caract";
                            parr1.Format.Alignment = ParagraphAlignment.Justify;
                            if (cumplimiento.EvaluacionID == 3)
                            {
                                parr1.Format.Font.Color = Colors.Blue;
                                if (Inspeccion.Fase > 1)
                                {
                                    var corregido = c.Cumplimiento
                                                        .Where(w => w.InspeccionID == Inspeccion.ID)
                                                        .FirstOrDefault();
                                    if (corregido != null)
                                    {
                                        parr1 = cRow.Cells[4].AddParagraph(string.Format("{0} en Fase {1}", corregido.Evaluacion.Descripcion, ToRoman(Inspeccion.Fase)));
                                        parr1.Style = "Carac";
                                        parr1.Format.Alignment = ParagraphAlignment.Justify;
                                        parr1.Format.Font.Color = Colors.Blue;
                                        parr1.Format.Shading.Color = Colors.Yellow;
                                    }

                                }

                            }
                        }
                        subsubpoint++;
                    }
                    subpoint++;
                }
            }
        }
    }
    public void ObservacionesNormativasYTecnicas()
    {
        Section section = document.AddSection();
        point++;
        subpoint = 1;
        Paragraph title = section.AddParagraph(string.Format("{0}. RESUMEN DE OBSERVACIONES NORMATIVAS Y TÉCNICAS", point));
        title.Style = "Heading1";
        title.AddBookmark("observaciones");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}. RESUMEN DE OBSERVACIONES NORMATIVAS Y TÉCNICAS", point), Mark = "observaciones", IsSub = false });

        Paragraph texto = section.AddParagraph(string.Format("Las siguientes observaciones deben ser corregidas para que el elevador quede en norma, y pueda ser certificado:", Inspeccion.Aparato.Nombre));
        texto.Style = "Parrafo";
        title = section.AddParagraph(string.Format("{0}.{1} OBSERVACIONES POR NORMA", point, subpoint));
        title.Style = "Heading2";
        title.AddBookmark("observacionespornorma");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}.{1} OBSERVACIONES POR NORMA", point, subpoint), Mark = "observacionespornorma", IsSub = true });
        // Observaciones por Norma
        var insp = Inspeccion.Fase == 1 ? Inspeccion : Inspeccion.Inspeccion2;
        var noCumplimiento = insp.Cumplimiento
                            .Where(w => w.EvaluacionID == 3 || w.EvaluacionID == 1)
                            .Where(w => w.EvaluacionID == 3 ? w.Observacion != null || w.Fotografias.Count > 0
                                    : w.Fotografias.Count > 0)
                            .Where(w => w.Caracteristica.Habilitado == true)
                            .Select(s => new
                            {
                                Requisito = s.Caracteristica.Requisito.Descripcion,
                                Norma = s.Caracteristica.Requisito.Titulo.Norma.Nombre,
                                Observacion = s.Observacion,
                                Fotos = s.Fotografias.Select(f => f.URL),
                                Evaluacion = s.EvaluacionID,
                                CaracteristicaId = s.CaracteristicaID
                            })
                            .OrderBy(o => o.Evaluacion)
                            .ThenBy(o => o.Fotos.Count() > 0)

                            ;
        if (!noCumplimiento.Any())
            return;

        var subsubpoint = 1;
        var numberfoto = 1;
        string pathImage = HttpContext.Current.Server.MapPath("~/fotos/");

        var noCumplimientoSinFoto = noCumplimiento.Where(w => !w.Fotos.Any());
        var noCumplimientoConFoto = noCumplimiento.Where(w => w.Fotos.Any());
        var count = 0;
        if (noCumplimientoSinFoto.Count() > 0)
        {
            foreach (var nc in noCumplimientoSinFoto)
            {
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                var complemento = nc.Evaluacion == 3
                                    ? string.Format("No cumple con el punto {0} de la norma {1}.", puntoNC, nc.Norma)
                                    : string.Empty;
                texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3} {4}", point, subpoint, subsubpoint, (nc.Observacion ?? string.Empty), complemento));
                texto.Style = "Parrafo";
                //texto.Format.Alignment = ParagraphAlignment.Left;
                if (Inspeccion.Fase == 2 && nc.Evaluacion == 3)
                {
                    var ok = Inspeccion.Cumplimiento
                                .Where(w => w.CaracteristicaID == nc.CaracteristicaId)
                                .FirstOrDefault();
                    if (ok != null)
                    {
                        texto = section.AddParagraph(ok.Evaluacion.Descripcion + " en Fase " + ToRoman(2));
                        texto.Style = "Parrafo";
                        texto.Format.Font.Color = Colors.Blue;
                    }
                }
                subsubpoint++;
            }
            section.AddPageBreak();
        }
        if (noCumplimientoConFoto.Count() > 0)
        {
            foreach (var nc in noCumplimientoConFoto)
            {
                if (count == 2)
                {
                    section.AddPageBreak();
                    count = 0;
                }
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                var complemento = nc.Evaluacion == 3
                                    ? string.Format("No cumple con el punto {0} de la norma {1}.", puntoNC, nc.Norma)
                                    : string.Empty;
                texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3} {4}", point, subpoint, subsubpoint, (nc.Observacion ?? string.Empty), complemento));
                texto.Style = "Parrafo";
                texto.Format.Alignment = ParagraphAlignment.Left;
                if (Inspeccion.Fase == 2 && nc.Evaluacion == 3)
                {
                    var ok = Inspeccion.Cumplimiento
                                .Where(w => w.CaracteristicaID == nc.CaracteristicaId)
                                .FirstOrDefault();
                    if (ok != null)
                    {
                        texto = section.AddParagraph(ok.Evaluacion.Descripcion + " en Fase " + ToRoman(2));
                        texto.Style = "Parrafo";
                        texto.Format.Font.Color = Colors.Blue;
                    }
                }
                subsubpoint++;

                foreach (var foto in nc.Fotos)
                {
                    var p = section.AddParagraph("");
                    p.Format.Alignment = ParagraphAlignment.Center;
                    Image image = section.LastParagraph.AddImage(pathImage + "/" + foto);
                    image.Width = "8cm";
                    var parr = section.AddParagraph("Imagen N° " + numberfoto);
                    parr.Style = "Pie";
                    numberfoto++;

                }
                count++;
            }
        }

        var observacionesTecnicas = insp.ObservacionTecnica;
        if (observacionesTecnicas.Count == 0)
            return;

        subpoint++;
        title = section.AddParagraph(string.Format("{0}.{1} OBSERVACIONES TÉCNICAS", point, subpoint));
        title.Style = "Heading1";
        title.AddBookmark("observacionestecnicas");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}.{1} OBSERVACIONES TÉCNICAS", point, subpoint), Mark = "observacionestecnicas", IsSub = true });

        subsubpoint = 1;
        count = 0;
        var otSinFoto = observacionesTecnicas.Where(a => !a.FotografiaTecnica.Any());
        var otConFoto = observacionesTecnicas.Where(a => a.FotografiaTecnica.Any());
        if (otSinFoto.Count() > 0)
        {
            foreach (var o in otSinFoto)
            {

                texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3}", point, subpoint, subsubpoint, (o.Texto ?? string.Empty)));
                texto.Style = "Parrafo";
                subsubpoint++;

                if(Inspeccion.Fase == 2)
                {
                    texto = section.AddParagraph(o.CorregidoEnFase2 == true ? "Corregido en Fase II" : "No corregido en Fase II");
                    texto.Style = "Parrafo";
                    texto.Format.Font.Color = Colors.AliceBlue;
                }
                

            }
            section.AddPageBreak();
        }
        if (otConFoto.Count() > 0)
        {
            foreach (var o in otConFoto)
            {
                if (count == 2)
                {
                    section.AddPageBreak();
                    count = 0;
                }
                texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3}", point, subpoint, subsubpoint, (o.Texto ?? string.Empty)));
                texto.Style = "Parrafo";
                subsubpoint++;
                if(Inspeccion.Fase == 2)
                {
                    texto = section.AddParagraph(o.CorregidoEnFase2 == true ? "Corregido en Fase II" : "No corregido en Fase II");
                    texto.Style = "Parrafo";
                    texto.Format.Font.Color = Colors.Blue;
                }
               
                var photo = o.FotografiaTecnica.Select(s => s.URL).FirstOrDefault();
                var p = section.AddParagraph("");
                p.Format.Alignment = ParagraphAlignment.Center;
                Image image = section.LastParagraph.AddImage(pathImage + "/" + photo);

                image.Width = "8cm";
                var parr = section.AddParagraph("Imagen N° " + numberfoto);
                parr.Style = "Pie";
                numberfoto++;
                count++;

            }
        }
    }
    public void Conclusiones()
    {
        Section section = document.AddSection();

        point++;
        subpoint = 1;
        Paragraph title = section.AddParagraph(string.Format("{0}. CONCLUSIONES", point));
        title.Style = "Heading1";
        title.AddBookmark("conclusiones");
        bookMarkList.Add(new BookMark { Text = string.Format("{0}. CONCLUSIONES", point), Mark = "conclusiones", IsSub = false });
        //Paragraph texto = section.AddParagraph(string.Format("Es necesario dar solución a las no conformidades y observaciones encontradas tras el proceso de inspección demoninado Fase {0}, separando las observaciones correspondientes a la edificación (cliente), así como las correspondientes a la empresa instaladora/mantenedora de ascensores,  con el objeto de incrementar la seguridad del mismo, proteger adecuadamente a los usuarios, a los técnicos de mantención, certificadores y/o personal propio del edificio en labores de rescate.", Inspeccion.Fase));
        //texto.Style = "Parrafo";
        //texto = section.AddParagraph(string.Format("Se debe trabajar en las mejoras de las no conformidades y observaciones normativas y técnicas descritas en los puntos 4 y 5 del presente informe, para que el {0} pueda calificar para la certificación sin observaciones y así, cumpla con la Ley 20.296.", Inspeccion.Aparato.Nombre));
        //texto.Style = "Parrafo";
        //texto = section.AddParagraph(string.Format("Es importante que tanto la administración del edificio, como la empresa instaladora/mantenedora, colaboran con la implementación de la carpeta cero, ya que existen en ella documentos que servirán para inscribir el {0} en la DOM (Dirección de Obras Municipales), según la indicación de la OGUC Artículo 5.9.5. Numeral 1, mediante una identificación con número único de registro de elevador.", Inspeccion.Aparato.Nombre));
        //texto.Style = "Parrafo";
        Paragraph texto;
        var tipoCalificacion = Inspeccion.Calificacion; // 0: no califica; 1: califica con observaciones menores; 2: califica sin observaciones
        var nAsociadas = Inspeccion.InspeccionNorma
                                .Select(s => s.Norma.Nombre)
                                .Distinct()
                                .ToArray();
        var normas = string.Empty;
        for(var i = 0; i < nAsociadas.Length; i++)
        {
            var isLast = i == nAsociadas.Length - 1;
            if(!isLast && nAsociadas.Length > 1)
            {
                normas += nAsociadas[i] + ", ";
            }
            else if(nAsociadas.Length > 1)
            {
                normas += "y " + nAsociadas[i];
            }
            else
            {
                normas += nAsociadas[i];
            }
        }
        switch(tipoCalificacion) 
        {
            case 0: // NO CALIFICA
                texto = section.AddParagraph(string.Format("Es necesario dar solución a las no conformidades y observaciones encontradas, separando las correspondientes a la edificación (cliente), así como las correspondientes a la empresa mantenedora de ascensores,  con el objeto de incrementar la seguridad del mismo, proteger adecuadamente a los usuarios, a los técnicos de mantención y/o personal propio de la empresa en labores de rescate de emergencia."));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("La OGUC (Ordenanza General de Urbanismo y Construcciones) en el Artículo 5.1.6, Numeral 13, indica que los elevadores deben disponer de una carpeta cero (o carpeta del elevador), este requisito es reafirmado por el punto Registros, de la norma {0} que indica la documentación necesaria que debe disponer dicha carpeta.", normas));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Es importante que tanto la administración del Edificio {0}, como la empresa mantenedora, colaboren en la implementación de la carpeta cero,  ya que existen en ella documentos que servirán para inscribir el ascensor en la DOM (Dirección de Obras Municipales) según la indicación de la OGUC Artículo 5.9.5. Numeral 1, mediante una identificación con número único de registro del elevador.", Inspeccion.NombreEdificio));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("La {0} N° {1}, en su estado actual, NO CALIFICA PARA LA CERTIFICACIÓN POR DEFECTOS GRAVES, según  las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas. Se recomienda  corregir las no conformidades y observaciones técnicas según la norma {2} señaladas en los puntos 4 y 5 del presente informe para que la {0} pueda cumplir con las normas Chilenas y pueda certificarse sin observaciones.", Inspeccion.Aparato.Nombre, Inspeccion.Numero, normas));
                texto.Style = "Parrafo";
                if(Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("Se da un plazo de {0} días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de no conformidades de la {1}.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString(), Inspeccion.Aparato.Nombre));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph("Cumplido este plazo, se programará en conjunto con el cliente, la Fase II del servicio,  para revisar si lo solicitado/sugerido en este informe, fue realizado, y así verificar si el equipo califica o no para su certificación.");
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Si pasados los {0} días, no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString()));
                    texto.Style = "Parrafo";
                }
                else if (Inspeccion.CreaFaseSiguiente == true)
                {
                    texto = section.AddParagraph("Debido a que las no conformidades no fueron subsanadas tras la inspección del servicio de Fase II, el cliente debe solicitar el servicio de inspección de Fase III y/o iniciar el proceso de certificación nuevamente.");
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph("Si elige el servicio de Fase III (cotización adicional), se otorga un plazo de 30 días para realizar las mejoras pendientes (observaciones menores que no afecten el normal funcionamiento del elevador). Si tras la inspección de Fase III no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.");
                    texto.Style = "Parrafo";
                }
                break;
            case 2: // CALIFICA CON OBSERVACIONES MENORES
                texto = section.AddParagraph(string.Format("Es necesario dar solución a las no conformidades y observaciones encontradas, separando las correspondientes a la edificación (cliente), así como las correspondientes a la empresa mantenedora de ascensores,  con el objeto de incrementar la seguridad del mismo, proteger adecuadamente a los usuarios, a los técnicos de mantención y/o personal propio de la empresa en labores de rescate de emergencia."));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("La OGUC (Ordenanza General de Urbanismo y Construcciones) en el Artículo 5.1.6, Numeral 13, indica que los elevadores deben disponer de una carpeta cero (o carpeta del elevador), este requisito es reafirmado por el punto Registros, de la norma {0} que indica la documentación necesaria que debe disponer dicha carpeta.", normas));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Es importante que tanto la administración del Edificio {0}, como la empresa mantenedora, colaboren en la implementación de la carpeta cero,  ya que existen en ella documentos que servirán para inscribir el ascensor en la DOM (Dirección de Obras Municipales) según la indicación de la OGUC Artículo 5.9.5. Numeral 1, mediante una identificación con número único de registro del elevador.", Inspeccion.NombreEdificio));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("La {0} N° {1}, en su estado actual, CALIFICA PARA LA CERTIFICACIÓN CON DEFECTOS LEVES, según  las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas. Se recomienda  corregir las no conformidades y observaciones técnicas según la norma {2} señaladas en los puntos 4 y 5 del presente informe para que la {0} pueda cumplir con las normas Chilenas y pueda certificarse sin observaciones.", Inspeccion.Aparato.Nombre, Inspeccion.Numero, normas));
                texto.Style = "Parrafo";
                if (Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("Se da un plazo de {0} días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de no conformidades de la {1}.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString(), Inspeccion.Aparato.Nombre));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph("Cumplido este plazo, se programará en conjunto con el cliente, la Fase II del servicio,  para revisar si lo solicitado/sugerido en este informe, fue realizado, y así verificar si el equipo califica o no para su certificación.");
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Si pasados los {0} días, no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString()));
                    texto.Style = "Parrafo";
                }
                else if (Inspeccion.CreaFaseSiguiente == true)
                {
                    texto = section.AddParagraph("Debido a que las no conformidades no fueron subsanadas tras la inspección del servicio de Fase II, el cliente debe solicitar el servicio de inspección de Fase III y/o iniciar el proceso de certificación nuevamente.");
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph("Si elige el servicio de Fase III (cotización adicional), se otorga un plazo de 30 días para realizar las mejoras pendientes (observaciones menores que no afecten el normal funcionamiento del elevador). Si tras la inspección de Fase III no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.");
                    texto.Style = "Parrafo";
                }
                break;
            case 1: // CALIFICA SIN OBSERVACIONES
                texto = section.AddParagraph(string.Format("En conformidad a las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas, se acredita mediante  inspección técnica y normativa, que la instalación de la {0} cumple con los requisitos de instalación y de las seguridades en conformidad con las normas {1} aplicadas. Por lo tanto, se acredita que el elevador ha sido adecuadamente mantenido y que se encuentra en condiciones de seguir funcionando.", Inspeccion.Aparato.Nombre, normas));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("La {0} N° {1}, CALIFICA PARA LA CERTIFICACIÓN, cumpliendo con la Ley 20.296.", Inspeccion.Aparato.Nombre, Inspeccion.Numero));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("El certificado de inspección técnica y normativa denominado Certificado de Inspección Electromecánico, deberá ser ingresado a la Dirección de Obras Municipales respectiva, por el propietario o por el administrador, según corresponda, antes del vencimiento del plazo que tiene la instalación para certificarse, y dentro de un plazo no superior a {0} días contados desde la fecha de emisión de la certificación.", "30"));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Se procederá entonces, a emitir el certificado de inspección electromecánico y de experiencia del elevador, el que estará disponible para su despacho en un plazo máximo de {0} días hábiles.", "5"));
                texto.Style = "Parrafo";
                
                
                break;
        }

        texto = section.AddParagraph("Atentamente,");
        texto.Style = "Parrafo";

        texto = section.AddParagraph("DEPARTAMENTO DE INGENIERÍA.");
        texto.Style = "Parrafo";
        texto.Format.Font.Bold = true;

        string pathImage = HttpContext.Current.Server.MapPath("~/css/images/");
        Image image = section.AddImage(pathImage + "/logo.png");
        image.Width = "5cm";
        image.Top = 10;

    }
    public void DefineTableOfContents(Document document)
    {

        Section section = new Section();
        Sections sections = document.Sections;

        Paragraph paragraph = section.AddParagraph("ÍNDICE");
        paragraph.Style = "Heading1";

        section.PageSetup.TopMargin = 180;
        foreach (var b in bookMarkList)
        {
            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            Hyperlink hyperlink = paragraph.AddHyperlink(b.Mark);
            var tab = b.IsSub ? " · " : "";
            hyperlink.AddText(tab + b.Text + "\t");
            hyperlink.AddPageRefField(b.Mark);
        }

        sections.InsertObject(1, section);
    }
    public string Rendering()
    {
        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
        pdfRenderer.Document = document;
        pdfRenderer.RenderDocument();
        var date = DateTime.Now.ToString("ddMMyyyyHHmmss");
        string filename = string.Format("Informe Inspeccion IT {0}_{1}.pdf", Inspeccion.IT.Replace("/", "-"), date);
        string basePath = HttpContext.Current.Server.MapPath("~/pdf/");
        string path = basePath + filename;

        using (var db = new CertelEntities())
        {
            var existsInforme = db.Informe
                                    .Where(w => w.InspeccionID == Inspeccion.ID)
                                    .FirstOrDefault();
            if (existsInforme == null)
            {
                var informe = new Informe
                {
                    FechaElaboracion = DateTime.Now,
                    EstadoID = 1,
                    InspeccionID = Inspeccion.ID,
                    FileName = filename,
                };
                db.Informe.Add(informe);
            }
            else
            {
                if (File.Exists(basePath + existsInforme.FileName))
                    File.Delete(basePath + existsInforme.FileName);

                existsInforme.FileName = filename;
                existsInforme.FechaElaboracion = DateTime.Now;
                existsInforme.EstadoID++;
            }
            db.SaveChanges();
        }
        pdfRenderer.PdfDocument.Save(path);
        
        return filename;
    }
    public struct BookMark
    {
        public string Text { get; set; }
        public string Mark { get; set; }
        public bool IsSub { get; set; }
    }
}