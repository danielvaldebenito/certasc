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
public class CreatePdfInforme
{
    public static Document document;
    public static Inspeccion Inspeccion { get; set; }
    public string FileName { get; set; }
    int point = 1;
    int subpoint = 1;
    int page = 1;
    int dg = 0;
    int dl = 0;
    int ok = 0;
    int na = 0;
    int total = 0;
    int corregido = 0;
    int nocorregido = 0;
    string normas = string.Empty;
    public int NormaPrincipal { get; set; }
    public string NormaPrincipalNombre { get; set; }
    public int TipoInforme { get; set; }
    public string Rendered { get; set; }
    public Color lightGray = new Color(238, 238, 238);


    public CreatePdfInforme(Inspeccion inspeccion, int normaPrincipal, string normaPrincipalNombre, int tipoInforme)
    {
        Inspeccion = inspeccion;
        NormaPrincipal = normaPrincipal;
        NormaPrincipalNombre = normaPrincipalNombre;
        TipoInforme = tipoInforme;
        FileName = "Inspeccion IT " + Inspeccion.IT.Replace('/', '-') + ".pdf";
        document = new Document();
        document.Info.Title = "Inspección";
        document.DefaultPageSetup.TopMargin = "2cm";
        document.DefaultPageSetup.LeftMargin = "2cm";
        document.DefaultPageSetup.RightMargin = "2cm";


        var nAsociadas = Inspeccion.InspeccionNorma
                                .Select(s => s.Norma.Nombre)
                                .Distinct()
                                .ToArray();
        
        for (var i = 0; i < nAsociadas.Length; i++)
        {
            var isLast = i == nAsociadas.Length - 1;
            if (!isLast && nAsociadas.Length > 1)
            {
                normas += nAsociadas[i] + ", ";
            }
            else if (nAsociadas.Length > 1)
            {
                normas += "y " + nAsociadas[i];
            }
            else
            {
                normas += nAsociadas[i];
            }
        }

        DefineStyles(document);
        DefineCover(document);
        Antecedentes();
        TerminosYDefiniciones();
        ResultadosInspeccion();
        Resumen();
        Conclusiones();
        Rendered = Rendering();
    }
    public static void DefineStyles(Document document)
    {
        
        // Get the predefined style Normal.
        var style = document.Styles["Normal"];
        // Because all styles are derived from Normal, the next line changes the
        // font of the whole document. Or, more exactly, it changes the font of
        // all styles and paragraphs that do not redefine the font.
        style.Font.Name = "Arial";
        // Heading1 to Heading9 are predefined styles with an outline level. An outline level
        // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks)
        // in PDF.
        style = document.Styles["Heading1"];
        style.Font.Size = 10;
        style.Font.Bold = true;
        style.Font.Color = Colors.Black;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.PageBreakBefore = false;
        style.ParagraphFormat.SpaceAfter = "1cm";
        style.ParagraphFormat.SpaceBefore = "2cm";
        style.ParagraphFormat.Borders.Visible = true;
        style.ParagraphFormat.Borders.Color = Colors.Gray;
        style.ParagraphFormat.Borders.Distance = 5;

        // Parrafo Normal
        style = document.Styles.AddStyle("Parrafo", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.SpaceBefore = "0.2cm";
        style.ParagraphFormat.SpaceAfter = "0.2cm";

        // Caract
        style = document.Styles.AddStyle("Caract", "Normal");
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.SpaceBefore = "0.2cm";
        style.ParagraphFormat.SpaceAfter = "0.2cm";
        // New Styles
        style = document.Styles.AddStyle("Portada", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 10;
        style.ParagraphFormat.SpaceBefore = "0.3cm";
        style.ParagraphFormat.SpaceAfter = "0.3cm";
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Footer", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Font.Color = Colors.Blue;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Antecedentes", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.SpaceBefore = "0.5cm";
        style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        style = document.Styles.AddStyle("Header", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.Font.Bold = true;
        
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
        parr.Format.SpaceBefore = "3cm";
        Image image = section.LastParagraph.AddImage(pathImage + "/logo.png");
        image.Width = "8cm";
        paragraph = section.AddParagraph(string.Format("INFORME DE AUDITORÍA E INSPECCIÓN DEL {0} {1}", Inspeccion.Aparato.Nombre.ToUpper(), Inspeccion.TipoFuncionamientoAparato.Descripcion.ToUpper()));
        paragraph.Style = "Portada";
        paragraph.Format.SpaceBefore = "5cm";
        paragraph = section.AddParagraph(Inspeccion.Servicio.Cliente.Nombre);
        paragraph.Style = "Portada";

        paragraph = section.AddParagraph(Inspeccion.NombreEdificio);
        paragraph.Style = "Portada";

        paragraph = section.AddParagraph(Inspeccion.Numero);
        paragraph.Style = "Portada";
        paragraph.Format.SpaceAfter = "7cm";

        // Pie de pagina
        
        var table = section.AddTable();
        //table.Borders.Visible = true;
        table.Format.Alignment = ParagraphAlignment.Center;
        table.AddColumn(120);
        table.AddColumn(250);
        table.AddColumn(120);
        var row = table.AddRow();
        
        var parrFooter = row.Cells[0].AddParagraph();
        row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
        var img = parrFooter.AddImage(pathImage + "/logo.png");
        img.Width = "3cm";
        parrFooter = row.Cells[1].AddParagraph("Certificación de Ascensores S.A.\nCalle Tabancura N° 1613 Dpto. 701 Block C. Vitacura – Santiago\nTelf. (+56) 232273961 Cel. (+56) 944821821\nEmail: contacto@certasc.cl\nwww.certasc.cl");
        parrFooter.Format.Font.Size = 8;
        parrFooter.Format.Font.Color = Colors.Gray;
        row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
        parrFooter.Format.Alignment = ParagraphAlignment.Center;

        var parr1 = row.Cells[2].AddParagraph("Documento controlado");
        parr1.Format.Font.Bold = true;
        parr1.Format.Font.Size = 8;
        var t = row.Cells[2].Elements.AddTable();
        row.Cells[2].VerticalAlignment = VerticalAlignment.Center;
        t.Borders.Visible = true;
        t.Borders.Color = Colors.Gray;
        t.Borders.Width = 1;
        t.AddColumn(60);
        t.AddColumn(60);
        var r = t.AddRow();
        parr1 = r.Cells[0].AddParagraph("VERSIÓN");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
        parr1 = r.Cells[1].AddParagraph("1.0");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
        r = t.AddRow();
        parr1 =  r.Cells[0].AddParagraph("Fecha Aprobación");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
        parr1 =  r.Cells[1].AddParagraph("01-03-2017");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
        r = t.AddRow();
        parr1 =  r.Cells[0].AddParagraph("Código");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
        parr1 =  r.Cells[1].AddParagraph("F-11");
        parr1.Format.Font.Size = 8;
        parr1.Format.Font.Color = Colors.Gray;
        parr1.Format.Alignment = ParagraphAlignment.Center;
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
    
    
    public void Antecedentes()
    {
        var section = document.AddSection();

        var parr = section.AddParagraph(string.Format("Señores: {0}", Inspeccion.Servicio.Cliente.Nombre));
        parr.Style = "Antecedentes";
        parr.Format.SpaceBefore = "2cm";
        parr = section.AddParagraph(string.Format("Cliente: {0}", Inspeccion.NombreEdificio));
        parr.Style = "Antecedentes";
        parr = section.AddParagraph("Presente");
        parr.Style = "Antecedentes";
        parr = section.AddParagraph(string.Format("De acuerdo con la inspección realizada el día {0} en el {1}, se envía informe {2}, con el resultado de la revisión técnica y normativa, detallando las no conformidades que deberán ser regularizadas para inicial el proceso de certificación.", 
            Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"),
            Inspeccion.NombreEdificio,
            Inspeccion.IT));
        parr.Style = "Antecedentes";
        parr = section.AddParagraph("Quedamos a su disposición y atentos a cualquier consulta.");
        parr.Style = "Antecedentes";
        parr = section.AddParagraph("Saluda atentamente,\nDpto.Técnico Ingeniería CertAsc S.A.");
        parr.Style = "Antecedentes";
        parr.Format.Alignment = ParagraphAlignment.Left;
        parr.Format.SpaceAfter = "3cm";

        Paragraph tableTitle = section.AddParagraph(string.Format("INFORME DE AUDITORÍA TÉCNICA E INSPECCIÓN DEL {0}", Inspeccion.Aparato.Nombre.ToUpper()));
        tableTitle.Style = "Heading1";
        Table table1 = section.AddTable();
        table1.Borders.Visible = true;
        table1.Borders.Color = Colors.LightGray;
        table1.AddColumn(150);
        table1.AddColumn(130);
        table1.AddColumn(200);
    
        Row row = table1.AddRow();
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;

        row.Cells[0].MergeRight = 2;
        Paragraph parrafo1 = row.Cells[0].AddParagraph(string.Format("CONTROL DE GESTIÓN"));
        row.Cells[0].Shading.Color = lightGray;
        row = table1.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row.Cells[0].AddParagraph("Fecha de emisión");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"));
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr = row.Cells[2].AddParagraph(string.Format("Emitido por: {0} {1}", Inspeccion.Usuario.Nombre, Inspeccion.Usuario.Apellido));
        parr.Format.Alignment = ParagraphAlignment.Justify;

        row = table1.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row.Cells[0].AddParagraph("Fecha de revisión");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaRevision.HasValue ? Inspeccion.FechaRevision.Value.ToString("dd-MM-yyyy") : string.Empty);
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr = row.Cells[2].AddParagraph(string.Format("Revisado por: {0} {1}", Inspeccion.Usuario1.Nombre, Inspeccion.Usuario1.Apellido));
        parr.Format.Alignment = ParagraphAlignment.Justify;

        row = table1.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row.Cells[0].AddParagraph("Fecha de aprobación");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaAprobacion.HasValue ? Inspeccion.FechaAprobacion.Value.ToString("dd-MM-yyyy"): string.Empty);
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr = row.Cells[2].AddParagraph(string.Format("Aprobado por por: {0} {1}", Inspeccion.Usuario2.Nombre, Inspeccion.Usuario2.Apellido));
        parr.Format.Alignment = ParagraphAlignment.Justify;

        row = table1.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row.Cells[0].AddParagraph("Fecha de entrega");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaEntrega.HasValue ? Inspeccion.FechaEntrega.Value.ToString("dd-MM-yyyy") : string.Empty);
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr = row.Cells[2].AddParagraph(string.Format("Nombre Cliente: {0}", Inspeccion.Destinatario));
        parr.Format.Alignment = ParagraphAlignment.Justify;


        parr = section.AddParagraph();
        parr.Format.SpaceAfter = "1cm";

        var table2 = section.AddTable();
        table2.Borders.Visible = true;
        table2.Borders.Width = 1;
        table2.Borders.Color = Colors.LightGray;
        
        table2.AddColumn(200);
        table2.AddColumn(280);
        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;
        parr = row.Cells[0].AddParagraph("DATOS BÁSICOS");
        row.Cells[0].Shading.Color = lightGray;
        row.Cells[0].MergeRight = 1;
        row.BottomPadding = 2;
        row.TopPadding = 2;
        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        parr = row.Cells[0].AddParagraph("Nº Informe");
        parr = row.Cells[1].AddParagraph(Inspeccion.Servicio.IT);
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        parr = row.Cells[0].AddParagraph("Dirección");
        parr = row.Cells[1].AddParagraph(Inspeccion.Ubicacion);
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        parr = row.Cells[0].AddParagraph("Nombre Inspector");
        parr = row.Cells[1].AddParagraph(Inspeccion.Usuario.Nombre + " " + Inspeccion.Usuario.Apellido);
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        parr = row.Cells[0].AddParagraph("Fecha de inspección");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"));
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        parr = row.Cells[0].AddParagraph("Etapa del proceso"); 
        parr = row.Cells[1].AddParagraph(string.Format("Etapa {0}", ToRoman(Inspeccion.Fase)));
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table2.AddRow();
        row.BottomPadding = 2;
        row.TopPadding = 2;
        var normas = Inspeccion.InspeccionNorma.Select(s => s.Norma.Nombre).ToArray();

        parr = row.Cells[0].AddParagraph("Norma aplicada");
        parr = row.Cells[1].AddParagraph(string.Join("; ", normas));
        row.BottomPadding = 2;
        row.TopPadding = 2;


        section.AddPageBreak();

        var table3 = section.AddTable();
        table3.AddColumn(200);
        table3.AddColumn(280);
        table3.Borders.Visible = true;
        table3.Borders.Color = Colors.LightGray;
        table3.Borders.Width = 1;

        row = table3.AddRow();
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;
        parr = row.Cells[0].AddParagraph("DATOS DEL ASCENSOR");
        row.Cells[0].Shading.Color = lightGray;
        row.Cells[0].MergeRight = 1;
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Nombre edificio");
        parr = row.Cells[1].AddParagraph(Inspeccion.NombreEdificio);
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Número único del elevador");
        parr = row.Cells[1].AddParagraph(Inspeccion.Numero);
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Equipo Nº");
        parr = row.Cells[1].AddParagraph(Inspeccion.Numero); // ?
        row.BottomPadding = 2;
        row.TopPadding = 2;

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Destino de uso del elevador");
        parr = row.Cells[1].AddParagraph(Inspeccion.DestinoProyecto.Descripcion);
        row.BottomPadding = 2;
        row.TopPadding = 2;


        // Tabla 3

        var insp = Inspeccion.Fase == 1 ? Inspeccion : Inspeccion.Inspeccion2;




        // Especificos Tabla 3
        var especificosT2 = insp.ValoresEspecificos.OrderBy(o => o.EspecificoID);
        foreach (var e in especificosT2)
        {
            row = table3.AddRow();
            row.Cells[0].AddParagraph(e.Especificos.Nombre);
            row.Cells[1].AddParagraph(e.Valor);
            row.TopPadding = 2;
            row.BottomPadding = 2;
        }

        // <br>
        parr = section.AddParagraph();
        parr.Format.SpaceAfter = "1cm";


        Table table4 = section.AddTable();
        table4.Borders.Visible = true;
        table4.KeepTogether = true;
        table4.Borders.Color = Colors.LightGray;
        table4.AddColumn(240);
        table4.AddColumn(240);
       
        Row row2 = table4.AddRow();
        row2.Format.Font.Bold = true;
        row2.Format.Alignment = ParagraphAlignment.Center;
        row2.VerticalAlignment = VerticalAlignment.Center;
        row2.TopPadding = 5;
        row2.BottomPadding = 5;
        row2.Cells[0].MergeRight = 1;

        parrafo1 = row2.Cells[0].AddParagraph("CARACTERÍSTICAS GENERALES");

        row2.Cells[0].Shading.Color = lightGray;
        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Nombre del Proyecto");
        row2.Cells[1].AddParagraph(insp.NombreProyecto ?? string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Destino del Proyecto");
        row2.Cells[1].AddParagraph(insp.DestinoProyectoID == null ? string.Empty : insp.DestinoProyecto.Descripcion);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Permiso Edificación");
        row2.Cells[1].AddParagraph(insp.PermisoEdificacion ?? "Sin información");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Recepción Municipal");
        row2.Cells[1].AddParagraph(insp.RecepcionMunicipal ?? "Sin información");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Número único del elevador");
        row2.Cells[1].AddParagraph(insp.Numero ?? string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Fecha de inicio del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph(insp.FechaEmisionCertificado.HasValue ? insp.FechaEmisionCertificado.Value.ToString("dd-MM-yyyy") : "En proceso de certificación");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table4.AddRow();
        row2.Cells[0].AddParagraph("Fecha de vencimiento del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph(insp.FechaVencimientoCertificado.HasValue ? insp.FechaVencimientoCertificado.Value.ToString("dd-MM-yyyy") ?? string.Empty : string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;
    }

    public void TerminosYDefiniciones()
    {
        Section section = document.LastSection;
        Paragraph title = section.AddParagraph("TÉRMINOS Y DEFINICIONES");
        title.Style = "Heading1";
        var table = section.AddTable();
        table.AddColumn(200);
        table.AddColumn(290);
        
        using (var db = new CertelEntities())
        {
            var terminos = db.TerminosYDefiniciones
                            .Where(w => w.NormaID == NormaPrincipal);
            foreach (var t in terminos)
            {
                var row = table.AddRow();
                var parr1 = row.Cells[0].AddParagraph(t.Termino);
                parr1.Format.Font.Bold = true;
                parr1.Format.Alignment = ParagraphAlignment.Left;
                row.Cells[1].AddParagraph(t.Definicion.TrimEnd());
            }
        }

        title = section.AddParagraph("CRITERIOS DE CALIFICACIÓN DE DEFECTOS SEGÚN D.S. N°08 (V y U)");
        title.Style = "Heading1";

        var parr = section.AddParagraph("D.S. N°08 (V y U) de fecha 28 de agosto de 2017, Modifica Decreto Supremo N° 47, de Vivienda y Urbanismo, de 1992, Ordenanza General de urbanismo y Construcciones en materia de Ascensores.");
        parr.Style = "Parrafo";
        parr = section.AddParagraph("Artículo 4°, Para los efectos de lo dispuesto en el numeral 4, del párrafo décimo cuarto del artículo 5.9.5. de la Ordenanza General de Urbanismo y Construcciones se actualizan para calificar los defectos encontrados en las instalaciones al momento de efectuar la inspección que antecede a la certificación, estos serán calificados como defectos graves y defectos leves.");
        parr.Style = "Parrafo";
        parr = section.AddParagraph("DEFECTO GRAVE: Es todo aquél que constituye un riesgo para la seguridad de las personas, del personal técnico que mantiene las respectivas instalaciones, o de la instalación propiamente tal.");
        parr.Style = "Parrafo";
        parr = section.AddParagraph("En virtud de lo anterior, será considerado como grave todo aquel defecto que altere o pueda alterar el correcto funcionamiento de cualquiera de los sistemas o componentes de la respectiva instalación, señalados a continuación, cuando pueda causar un accidente por cizallamiento, aplastamiento, caída, choque, atrapamiento, fuego o choque eléctrico: ");
        parr.Style = "Parrafo";
        parr = section.AddParagraph("• Sistema de apertura de puertas, contactos de seguridad y dispositivos de enclavamiento.\n• Conjunto limitador de velocidad y paracaídas del equipo.\n• Sistemas de frenos del equipo.\n• Sistemas de suspensión y polea motriz, en especial cuando estos no cumplan con las disposiciones de seguridad especificadas por el fabricante.\n• Línea eléctrica o circuito de seguridad, incluidos los dispositivos de final de recorrido.\n• Registros carpeta de ascensores.");
        parr.Style = "Parrafo";
        parr.Format.Alignment = ParagraphAlignment.Left;
        parr = section.AddParagraph("DEFECTO LEVE: Es todo aquel no calificable como grave, y que por sí solo no significa un riesgo para la seguridad de las personas, para el personal técnico que mantiene las respectivas instalaciones, o para la instalación propiamente tal.");
        parr.Style = "Parrafo";
        parr = section.AddParagraph("En caso de que, conforme a las normas técnicas oficiales vigentes aplicables a la respectiva instalación, haya razones técnicas por las cuales estos defectos no puedan ser subsanados, el certificador deberá determinar fundadamente una solución alternativa para cada defecto, de carácter permanente, así como el plazo de ejecución de la misma solución, lo que deberá quedar detallado en un informe de defectos leves que se adjuntará a la certificación.");
        parr.Style = "Parrafo";
    }
    public void ResultadosInspeccion()
    {
        Section section = document.LastSection;
        
        using (var db = new CertelEntities())
        {
            var parr = section.AddParagraph("TERMINOLOGÍA: Defecto grave (DG), Defecto leve, (DL), No aplica (N/A), Cumple con el requisito (OK).");
            parr.Format.Font.Bold = true;
            parr.Format.Font.Size = 9;
            parr.Format.SpaceAfter = 10;
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

                Table table = section.AddTable();
                table.Borders.Visible = true;
                table.Borders.Color = Colors.Gray;
                table.KeepTogether = false;
                
                if (Inspeccion.Fase == 1)
                {
                    table.AddColumn(160);
                    table.AddColumn(166);
                    table.AddColumn(40);
                    table.AddColumn(40);
                    table.AddColumn(40);
                    table.AddColumn(40);
                }
                else
                {
                    table.AddColumn(160);
                    table.AddColumn(286);
                    table.AddColumn(40);
                }
                

                // TITULO
                Row row = table.AddRow();
                row.TopPadding = 5;
                row.BottomPadding = 5;
                row.Format.Font.Bold = true;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = VerticalAlignment.Center;
                row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 2;
                
                var title = row.Cells[0].AddParagraph(t.Texto);
                row.BottomPadding = 5;
                row.TopPadding = 5;
                // ENCABEZADO
                row = table.AddRow();
                row.Shading.Color = lightGray;
                row.Cells[0].MergeDown = 1;
                var header = row.Cells[0].AddParagraph(string.Format("{0}", n.Nombre));
                header.Style = "Header";
                header.Format.Alignment = ParagraphAlignment.Center;
                row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row.Cells[1].MergeDown = 1;
                header = row.Cells[1].AddParagraph("Criterio de aceptación");
                header.Style = "Header";
                header.Format.Alignment = ParagraphAlignment.Center;
                row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
                header = row.Cells[2].AddParagraph("Observaciones");
                header.Style = "Header";
                header.Format.Alignment = ParagraphAlignment.Center;
                row.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                row.Cells[2].VerticalAlignment = VerticalAlignment.Center;

                // SUB-ENCABEZADO
                row = table.AddRow();
                row.Shading.Color = lightGray;
                if (Inspeccion.Fase == 1)
                {
                    header = row.Cells[2].AddParagraph("OK");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;
                    header = row.Cells[3].AddParagraph("DG");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;
                    header = row.Cells[4].AddParagraph("DL");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;
                    header = row.Cells[5].AddParagraph("N/A");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;

                    row.Cells[4].VerticalAlignment = VerticalAlignment.Center;
                    row.Cells[5].VerticalAlignment = VerticalAlignment.Center;
                    
                }
                else
                {
                    header = row.Cells[2].AddParagraph("Corregido");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;
                    header = row.Cells[3].AddParagraph("No Corregido");
                    header.Style = "Header";
                    header.Format.Alignment = ParagraphAlignment.Center;

                }
                row.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                row.Cells[3].VerticalAlignment = VerticalAlignment.Center;




                var requisitos = t.Requisito.Where(w => w.Habilitado == true);
                var rIndex = 0;
                foreach (var r in requisitos)
                {
                    var cars = r.Caracteristica.Where(w => w.Habilitado == true);
                    var carsCount = cars.Count();
                    if (carsCount == 0)
                        continue;

                    rIndex++;

                    foreach (var c in cars)
                    {
                        
                        var cRow = table.AddRow();
                        if (rIndex % 2 != 0)
                        {
                            cRow.Shading.Color = lightGray;
                        }
                        cRow.Format.Alignment = ParagraphAlignment.Center;
                        cRow.VerticalAlignment = VerticalAlignment.Center;
                        cRow.TopPadding = 0;
                        cRow.BottomPadding = 0;
                        var parr1 = cRow.Cells[1].AddParagraph(c.Descripcion);
                        parr1.Style = "Caract";
                        parr1 = cRow.Cells[0].AddParagraph(string.Format("{0}", r.Descripcion));
                        parr1.Style = "Caract";
                        cRow.Cells[0].MergeDown = carsCount - 1;

                        var cumplimiento = c.Cumplimiento
                                            .Where(w => Inspeccion.Fase == 1
                                                    ? w.InspeccionID == Inspeccion.ID
                                                    : w.InspeccionID == Inspeccion.InspeccionFase1)
                                            .FirstOrDefault();
                        if (cumplimiento == null)
                            continue;

                        var index = 0;
                        total++;
                        switch (cumplimiento.EvaluacionID) {
                            case 1: ok++; index = 2; break;
                            case 2: dl++; index = 4; break;
                            case 3: dg++; index = 3; break;
                            case 4: na++; index = 5; break;
                            case 5: corregido++;  index = 2; break;
                            case 6: nocorregido++;  index = 3; break;
                        }
                        cRow.Cells[index].VerticalAlignment = VerticalAlignment.Center;
                        parr1 = cRow.Cells[index].AddParagraph(cumplimiento == null ? string.Empty : cumplimiento.Evaluacion.Glosa);
                        parr1.Format.Font.Size = 9;
                        parr1.Format.Font.Color = Colors.Black;
                        parr1.Format.Alignment = ParagraphAlignment.Center;

                        if (cumplimiento.EvaluacionID == 3) // defecto grave
                        {
                            parr1.Format.Font.Color = Colors.Gray;
                            if (Inspeccion.Fase > 1)
                            {
                                var corregido = c.Cumplimiento
                                                    .Where(w => w.InspeccionID == Inspeccion.ID)
                                                    .FirstOrDefault();
                                if (corregido != null)
                                {
                                    parr1 = cRow.Cells[3].AddParagraph(cumplimiento.Evaluacion.Glosa);
                                    parr1.Format.Font.Size = 9;
                                    parr1.Format.Font.Color = Colors.Black;
                                    parr1.Format.Alignment = ParagraphAlignment.Center;
                                }

                            }

                        }


                    }

                }

                section.AddParagraph();

            }
            var normasAsociadas = n.NormasAsociadas;

            foreach (var nor in normasAsociadas)
            {
                var ns = db.Norma.Find(nor.NormaSecundariaID);
                var titles = ns.Titulo.ToList();
                foreach (var t in titles)
                {

                    Table table = section.AddTable();
                    table.Borders.Visible = true;
                    table.Borders.Color = Colors.Gray;
                    table.Format.KeepTogether = false;
                    if (Inspeccion.Fase == 1)
                    {
                        table.AddColumn(160);
                        table.AddColumn(166);
                        table.AddColumn(40);
                        table.AddColumn(40);
                        table.AddColumn(40);
                        table.AddColumn(40);
                    }
                    else
                    {
                        table.AddColumn(160);
                        table.AddColumn(286);
                        table.AddColumn(40);
                    }
                    Row row = table.AddRow();
                    row.Format.Font.Bold = true;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = VerticalAlignment.Center;
                    row.TopPadding = 5;
                    row.BottomPadding = 5;
                    row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;
                    row.Cells[0].AddParagraph(string.Format("{0}", ns.TituloRegulacion));

                    row = table.AddRow();
                    var header = row.Cells[0].AddParagraph(ns.Nombre);
                    header.Style = "Header";
                    header = row.Cells[1].AddParagraph("Criterio de aceptación");
                    header.Style = "Header";
                    header = row.Cells[2].AddParagraph("Observaciones");
                    header.Style = "Header";
                    row.Cells[3].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;


                    row = table.AddRow();
                    row.Format.Font.Bold = true;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = VerticalAlignment.Center;
                    if (Inspeccion.Fase == 1)
                    {
                        header = row.Cells[2].AddParagraph("OK");
                        header.Style = "Header";
                        header = row.Cells[3].AddParagraph("DG");
                        header.Style = "Header";
                        header = row.Cells[4].AddParagraph("DL");
                        header.Style = "Header";
                        header = row.Cells[5].AddParagraph("N/A");
                        header.Style = "Header";
                    }
                    else
                    {
                        header = row.Cells[2].AddParagraph("Corregido");
                        header.Style = "Header";
                        header = row.Cells[3].AddParagraph("No Corregido");
                        header.Style = "Header";
                    }

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
                            var parr1 = cRow.Cells[1].AddParagraph(c.Descripcion);
                            parr1.Style = "Caract";
                            cRow.Cells[0].AddParagraph(string.Format("{0}", r.Descripcion));


                            cRow.Cells[0].MergeDown = carsCount - 1;

                            var cumplimiento = c.Cumplimiento
                                            .Where(w => Inspeccion.Fase == 1 ? w.InspeccionID == Inspeccion.ID : w.InspeccionID == Inspeccion.InspeccionFase1).FirstOrDefault();
                            if (cumplimiento == null)
                                continue;

                            var index = 0;
                            switch (cumplimiento.EvaluacionID)
                            {
                                case 1: ok++; index = 2; break;
                                case 2: dl++; index = 4; break;
                                case 3: dg++; index = 3; break;
                                case 4: na++; index = 5; break;
                                case 5: corregido++; index = 2; break;
                                case 6: nocorregido++; index = 3; break;
                            }
                            parr1 = cRow.Cells[index].AddParagraph(cumplimiento == null ? string.Empty : cumplimiento.Evaluacion.Glosa);

                            parr1.Format.Font.Size = 9;
                            parr1.Format.Font.Color = Colors.Black;
                            parr1.Format.Alignment = ParagraphAlignment.Center;

                            if (cumplimiento.EvaluacionID == 3)
                            {
                                parr1.Format.Font.Color = Colors.Blue;
                                if (Inspeccion.Fase > 1)
                                {
                                    var corregido = c.Cumplimiento
                                                        .Where(w => w.InspeccionID == Inspeccion.ID)
                                                        .FirstOrDefault();
                                    parr1 = cRow.Cells[corregido != null ? 3 : 4].AddParagraph(cumplimiento.Evaluacion.Glosa);
                                    parr1.Format.Font.Size = 9;
                                    parr1.Format.Font.Color = Colors.Black;
                                    parr1.Format.Alignment = ParagraphAlignment.Center;
                                }

                            }
                        }
                    }
                }
            }
        }
    }
    public void Conclusiones()
    {
        Section section = document.AddSection();

        point++;
        subpoint = 1;
        Paragraph title = section.AddParagraph("CONCLUSIONES");
        title.Style = "Heading1";
        //Paragraph texto = section.AddParagraph(string.Format("Es necesario dar solución a las no conformidades y observaciones encontradas tras el proceso de inspección demoninado Fase {0}, separando las observaciones correspondientes a la edificación (cliente), así como las correspondientes a la empresa instaladora/mantenedora de ascensores,  con el objeto de incrementar la seguridad del mismo, proteger adecuadamente a los usuarios, a los técnicos de mantención, certificadores y/o personal propio del edificio en labores de rescate.", Inspeccion.Fase));
        //texto.Style = "Parrafo";
        //texto = section.AddParagraph(string.Format("Se debe trabajar en las mejoras de las no conformidades y observaciones normativas y técnicas descritas en los puntos 4 y 5 del presente informe, para que el {0} pueda calificar para la certificación sin observaciones y así, cumpla con la Ley 20.296.", Inspeccion.Aparato.Nombre));
        //texto.Style = "Parrafo";
        //texto = section.AddParagraph(string.Format("Es importante que tanto la administración del edificio, como la empresa instaladora/mantenedora, colaboran con la implementación de la carpeta cero, ya que existen en ella documentos que servirán para inscribir el {0} en la DOM (Dirección de Obras Municipales), según la indicación de la OGUC Artículo 5.9.5. Numeral 1, mediante una identificación con número único de registro de elevador.", Inspeccion.Aparato.Nombre));
        //texto.Style = "Parrafo";
        Paragraph texto;
        var tipoCalificacion = Inspeccion.Calificacion; // 0: no califica; 1: califica con observaciones menores; 2: califica sin observaciones
        
        switch(tipoCalificacion) 
        {
            case 0: // NO CALIFICA
                texto = section.AddParagraph(string.Format("Según la evaluación del {0} {1}, se encuentran N° {2} de hallazgos denominados “Defectos graves” (DF) que corresponde al {3}% de la calificación y que también se encuentran N° {4} de hallazgos denominados “Defectos leves” (DL) correspondiente al {5}% por lo que de conformidad se encuentran N° {6} de “Conformidades” (OK) correspondiente al {7}% de un total de N° {8} de requisitos aplicados normativamente.", 
                    Inspeccion.Aparato.Nombre,
                    Inspeccion.TipoFuncionamientoAparato.Descripcion,
                    dg,
                    (dg * 100 / total).ToString(),
                    dl,
                    (dl * 100 / total).ToString(),
                    ok,
                    (ok * 100 / total).ToString(),
                    total
                ));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Según los datos estadísticos arrojados en la calificación, el {0} {1} N° {2} en su estado actual, no califica para la certificación por defectos leves y graves según las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 - D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas.", Inspeccion.Aparato.Nombre, Inspeccion.TipoFuncionamientoAparato.Descripcion, Inspeccion.Numero));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Se debe corregir los defectos graves y leves registrados en la inspección según las exigencias de las normas {0} señaladas en el presente informe para que el {1} pueda cumplir con las normas chilenas y certificarse sin observaciones.", normas, Inspeccion.Aparato.Nombre));
                texto.Style = "Parrafo";
                if(Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("Se otorga un plazo de {0} días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de hallazgos encontrados", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString()));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph("Cumplido este plazo, se programará en conjunto con el cliente, la Etapa II del servicio,  para revisar si lo solicitado/sugerido en este informe, fue realizado, y así verificar si el equipo califica o no para su certificación.");
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Si pasados los {0} días, no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString()));
                    texto.Style = "Parrafo";
                }
                break;
            case 2: // CALIFICA CON OBSERVACIONES MENORES
                texto = section.AddParagraph(string.Format("Según la evaluación del {0}, se encuentran N° {1} de hallazgos denominados “Defectos Leves” (DL) correspondiente al {2}% y de {3} de “Conformidades” (OK) correspondiente a {4}% de un total de {5} de requisitos aplicados normativamente.", Inspeccion.Aparato.Nombre, dl, Math.Round((double)(dl * 100 / total)), ok, Math.Round((double)(ok * 100 / total)), total));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("El {0} N°{1} en su estado actual, califica para la certificación con defectos leves, según las disposiciones contenidas en la Ley 20.296 y el D.S.N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S.N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas. ", Inspeccion.Aparato.Nombre, Inspeccion.Numero));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Se deben corregir los defectos leves y observaciones técnicas registradas en este informe según exigencias de la normas {0} señaladas en el presente informe para que el(ascensor o montacargas) pueda cumplir con las normas chilenas y certificarse sin observaciones.", normas));
                texto.Style = "Parrafo";
                
                texto.Style = "Parrafo";
                if (Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("Se otorga un plazo de {0} días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de los hallazgos encontrados.", Inspeccion.DiasPlazo ?? 90));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Cumplido este plazo, se programará en conjunto con el cliente la Etapa II del servicio, para revisar si lo solicitado/sugerido en este informe, fue realizado, y así verificar si el equipo califica o no para su certificación sin observaciones."));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Si pasados los {0} días, no se han realizado las mejoras; entonces se deberá comenzar nuevamente con el proceso de certificación; materia de otra cotización.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString()));
                    texto.Style = "Parrafo";
                }
                
                break;
            case 1: // CALIFICA SIN OBSERVACIONES
                texto = section.AddParagraph(string.Format("En conformidad a las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas, se acredita mediante inspección técnica y normativa, que la instalación del {0}, cumple con los requisitos de instalación y de las seguridades en conformidad con las normas {1} aplicadas. Por lo tanto, se acredita que el elevador ha sido adecuadamente mantenido y que se encuentran en condiciones de seguir funcionando.", Inspeccion.Aparato.Nombre, normas));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("El {1} N° {0} califica para la certificación, cumpliendo con la Ley 20.296. El certificado de inspección técnica y normativa denominado Certificado de Inspección Electromecánico, deberá ser ingresado a la Dirección de Obras Municipales respectiva, por el propietario o por el administrador, según corresponda, antes del vencimiento del plazo que tiene la instalación para certificarse, y dentro de un plazo no superior a 30 días contados desde la fecha de emisión de la certificación. Se procederá entonces, a emitir el certificado de inspección electromecánico y de experiencia del elevador, el que estará disponible para su despacho en un plazo máximo de 5 días hábiles.", Inspeccion.Numero, Inspeccion.Aparato.Nombre));
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
    public void Resumen()
    {
        var section = document.LastSection;
        var table = section.AddTable();
        table.Borders.Visible = true;
        table.Borders.Width = 1;
        table.Borders.Color = Colors.LightGray;
        if (Inspeccion.Fase == 1)
        {
            table.AddColumn(122);
            table.AddColumn(172);
            table.AddColumn(49);
            table.AddColumn(49);
            table.AddColumn(49);
            table.AddColumn(49);
        }
        else
        {
            table.AddColumn(122);
            table.AddColumn(172);
            table.AddColumn(98);
            table.AddColumn(98);
        }
        

        var row = table.AddRow();
        var parr = row.Cells[0].AddParagraph("RESUMEN ESTADÍSTICO DE CUMPLIMIENTOS");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 10;
        row.Cells[0].MergeRight = 5;
        row.BottomPadding = 5;
        row.TopPadding = 5;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("TOTAL REQUISITOS NORMATIVOS");
        parr.Format.Alignment = ParagraphAlignment.Right;
        parr.Format.Font.Size = 10;
        parr = row.Cells[1].AddParagraph(total.ToString());
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 10;
        parr = row.Cells[2].AddParagraph("CUMPLIMIENTOS");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 10;
        
        row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
        row.BottomPadding = 5;
        row.TopPadding = 5;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("TOTAL REQUISITOS APLICADOS");
        parr.Format.Alignment = ParagraphAlignment.Right;
        parr.Format.Font.Size = 10;
        parr = row.Cells[1].AddParagraph((total - na).ToString());
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 10;
        row.BottomPadding = 5;
        row.TopPadding = 5;

        if (Inspeccion.Fase == 1)
        {
            parr = row.Cells[2].AddParagraph("OK");
            
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[3].AddParagraph("DG");
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[4].AddParagraph("DL");
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[5].AddParagraph("N/A");
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            row.Cells[4].Shading.Color = lightGray;
            row.Cells[5].Shading.Color = lightGray;
        }
        else
        {
            parr = row.Cells[2].AddParagraph("Corregido");
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[3].AddParagraph("No Corregido");
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            
        }
        row.Cells[2].Shading.Color = lightGray;
        row.Cells[3].Shading.Color = lightGray;
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("CANTIDAD POR CALIFICACIÓN");
        row.Cells[0].MergeRight = 1;
        parr.Format.Alignment = ParagraphAlignment.Right;
        parr.Format.Font.Size = 9;
        

        if (Inspeccion.Fase == 1) // FIXME
        {
            parr = row.Cells[2].AddParagraph(ok.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[3].AddParagraph(dg.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[4].AddParagraph(dl.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[5].AddParagraph(na.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
        }
        else
        {
            parr = row.Cells[2].AddParagraph(corregido.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr = row.Cells[3].AddParagraph(corregido.ToString());
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;

        }

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("PORCENTAJE POR CALIFICACIÓN");
        parr.Format.Alignment = ParagraphAlignment.Right;
        parr.Format.Font.Size = 9;
        row.Cells[0].MergeRight = 1;

        if (Inspeccion.Fase == 1) // FIXME
        {
            parr = row.Cells[2].AddParagraph(String.Format("{0:.##}", (ok * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
            parr = row.Cells[3].AddParagraph(String.Format("{0:.##}", (dg * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
            parr = row.Cells[4].AddParagraph(String.Format("{0:.##}", (dl * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
            parr = row.Cells[5].AddParagraph(String.Format("{0:.##}", (na * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
        }
        else
        {
            parr = row.Cells[2].AddParagraph(String.Format("{0:.##}", (corregido * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
            parr = row.Cells[3].AddParagraph(String.Format("{0:.##}", (nocorregido * 100 / total)));
            parr.Format.Alignment = ParagraphAlignment.Center;
            parr.Format.Font.Size = 9;
            parr.Format.Font.Bold = true;
        }

        // fila vacia
        row = table.AddRow();
        row.Cells[0].AddParagraph(string.Empty);
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("RESUMEN DE OBSERVACIONES NORMATIVAS Y TÉCNICAS");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;
        row.BottomPadding = 5;
        row.TopPadding = 5;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Las siguientes observaciones deben ser corregidas para que el elevador quede en norma, y pueda ser certificado:");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;
        row.BottomPadding = 15;
        row.TopPadding = 15;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("OBSERVACIONES POR NORMA");
        parr.Format.Alignment = ParagraphAlignment.Center;
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;
        row.Cells[0].Shading.Color = lightGray;
        row.BottomPadding = 5;
        row.TopPadding = 5;

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("REQUISITO");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        parr = row.Cells[1].AddParagraph("DESCRIPCIÓN");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        parr = row.Cells[2].AddParagraph("IMAGEN");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
        row.BottomPadding = 3;
        row.TopPadding = 3;
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


        string pathImage = HttpContext.Current.Server.MapPath("~/fotos/");

        var noCumplimientoSinFoto = noCumplimiento.Where(w => !w.Fotos.Any());
        var noCumplimientoConFoto = noCumplimiento.Where(w => w.Fotos.Any());
        var count = 0;
        if (noCumplimientoSinFoto.Count() > 0)
        {
            var indexncsf = 0; // Indice no cumplimiento sin foto
            foreach (var nc in noCumplimientoSinFoto)
            {
                indexncsf++;
                
                row = table.AddRow();
                if (indexncsf % 2 != 0)
                {
                    row.Shading.Color = lightGray;
                }
                row.BottomPadding = 10;
                row.TopPadding = 10;
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                parr = row.Cells[0].AddParagraph(puntoNC);
                parr.Style = "Parrafo";
                parr = row.Cells[1].AddParagraph(nc.Observacion ?? string.Empty);
                parr.Style = "Parrafo";
                parr = row.Cells[2].AddParagraph();
                //texto.Format.Alignment = ParagraphAlignment.Left;
                //if (Inspeccion.Fase == 2 && nc.Evaluacion == 3)
                //{
                //    var ok = Inspeccion.Cumplimiento
                //                .Where(w => w.CaracteristicaID == nc.CaracteristicaId)
                //                .FirstOrDefault();
                //    if (ok != null)
                //    {
                //        parr = section.AddParagraph(ok.Evaluacion.Descripcion + " en Fase " + ToRoman(2));
                //        texto.Style = "Parrafo";
                //        texto.Format.Font.Color = Colors.Blue;
                //    }
            }
            section.AddParagraph();
        }
        if (noCumplimientoConFoto.Count() > 0)
        {
            var indexnccf = 0;
            foreach (var nc in noCumplimientoConFoto)
            {
                row = table.AddRow();

                if (indexnccf % 2 != 0)
                {
                    row.Shading.Color = lightGray;
                }
    
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                parr = row.Cells[0].AddParagraph(puntoNC);
                parr.Style = "Parrafo";
                parr = row.Cells[1].AddParagraph(nc.Observacion ?? string.Empty);
                parr.Style = "Parrafo";
                parr = row.Cells[2].AddParagraph();
                parr.Format.Alignment = ParagraphAlignment.Center;
                row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;

                //if (Inspeccion.Fase == 2 && nc.Evaluacion == 3)
                //{
                //    var ok = Inspeccion.Cumplimiento
                //                .Where(w => w.CaracteristicaID == nc.CaracteristicaId)
                //                .FirstOrDefault();
                //    if (ok != null)
                //    {
                //        texto = section.AddParagraph(ok.Evaluacion.Descripcion + " en Fase " + ToRoman(2));
                //        texto.Style = "Parrafo";
                //        texto.Format.Font.Color = Colors.Blue;
                //    }
                //}
                
                foreach (var foto in nc.Fotos)
                {
                    
                    Image image = parr.AddImage(pathImage + "/" + foto);
                    image.Width = 196;

                }
                count++;
            }
        }

        var observacionesTecnicas = insp.ObservacionTecnica;
        if (observacionesTecnicas.Count == 0)
            return;

        row = table.AddRow();
        row.Borders.Visible = false;
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("OBSERVACIONES TÉCNICAS");
        parr.Format.Alignment = ParagraphAlignment.Center;
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;
        row.BottomPadding = 5;
        row.TopPadding = 5;
        row.Cells[0].Shading.Color = lightGray;
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("REQUISITO");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        parr = row.Cells[1].AddParagraph("DESCRIPCIÓN");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        parr = row.Cells[2].AddParagraph("IMAGEN");
        parr.Format.Alignment = ParagraphAlignment.Center;
        parr.Format.Font.Size = 9;
        row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
        row.BottomPadding = 3;
        row.TopPadding = 3;

        var indexot = 0;
        foreach (var o in observacionesTecnicas)
        {
            row = table.AddRow();
            if (indexot % 2 != 0)
            {
                row.Shading.Color = lightGray;
            }

            parr = row.Cells[0].AddParagraph(string.Empty);
            parr.Style = "Parrafo";
            parr = row.Cells[1].AddParagraph(o.Texto);
            parr.Style = "Parrafo";
            parr = row.Cells[2].AddParagraph();
            parr.Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
            indexot++;
            var foto = o.FotografiaTecnica.FirstOrDefault();
            if (foto != null)
            {
                var img = parr.AddImage(pathImage + "/" + foto.URL);
                img.Width = 196;
            }
            
        }
        //subpoint++;
        //title = section.AddParagraph(string.Format("{0}.{1} OBSERVACIONES TÉCNICAS", point, subpoint));
        //title.Style = "Heading1";
        //title.AddBookmark("observacionestecnicas");

        //subsubpoint = 1;
        //count = 0;
        //var otSinFoto = observacionesTecnicas.Where(a => !a.FotografiaTecnica.Any());
        //var otConFoto = observacionesTecnicas.Where(a => a.FotografiaTecnica.Any());
        //if (otSinFoto.Count() > 0)
        //{
        //    foreach (var o in otSinFoto)
        //    {

        //        texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3}", point, subpoint, subsubpoint, (o.Texto ?? string.Empty)));
        //        texto.Style = "Parrafo";
        //        subsubpoint++;

        //        if (Inspeccion.Fase == 2)
        //        {
        //            texto = section.AddParagraph(o.CorregidoEnFase2 == true ? "Corregido en Fase II" : "No corregido en Fase II");
        //            texto.Style = "Parrafo";
        //            texto.Format.Font.Color = Colors.AliceBlue;
        //        }
        //    }
        //    section.AddPageBreak();
        //}
        //if (otConFoto.Count() > 0)
        //{
        //    foreach (var o in otConFoto)
        //    {
        //        if (count == 2)
        //        {
        //            section.AddPageBreak();
        //            count = 0;
        //        }
        //        texto = section.AddParagraph(string.Format("{0}.{1}.{2}. \t{3}", point, subpoint, subsubpoint, (o.Texto ?? string.Empty)));
        //        texto.Style = "Parrafo";
        //        subsubpoint++;
        //        if (Inspeccion.Fase == 2)
        //        {
        //            texto = section.AddParagraph(o.CorregidoEnFase2 == true ? "Corregido en Fase II" : "No corregido en Fase II");
        //            texto.Style = "Parrafo";
        //            texto.Format.Font.Color = Colors.Blue;
        //        }

        //        var photo = o.FotografiaTecnica.Select(s => s.URL).FirstOrDefault();
        //        var p = section.AddParagraph("");
        //        p.Format.Alignment = ParagraphAlignment.Center;
        //        Image image = section.LastParagraph.AddImage(pathImage + "/" + photo);

        //        image.Width = "8cm";
        //        var parr = section.AddParagraph("Imagen N° " + numberfoto);
        //        parr.Style = "Pie";
        //        numberfoto++;
        //        count++;

        //    }
        //}

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

}