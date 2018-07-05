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
public class CreatePDFD116
{
    public static Document document;
    public static Inspeccion Inspeccion { get; set; }
    public string FileName { get; set; }
    int point = 1;
    int subpoint = 1;
    int page = 1;
    public static int NormaPrincipal = 10;
    public static string NormaPrincipalNombre = "NCh440/1:2014";
    public static int TipoInforme = 3;
    public string Rendered { get; set; }
    public CreatePDFD116(Inspeccion inspeccion)
    {
        Inspeccion = inspeccion;
        FileName = "Inspeccion IT " + Inspeccion.IT.Replace('/', '-') + ".pdf";
        document = new Document();
        document.Info.Title = "Inspección";
        document.DefaultPageSetup.TopMargin = "2cm";
        document.DefaultPageSetup.LeftMargin = "2cm";
        document.DefaultPageSetup.RightMargin = "2cm";
        DefineStyles(document);
        DefineCover(document);
        Antecedentes();
        TerminosYDefiniciones();
        ResultadosInspeccion();
        Resumen();
        ObservacionesNormativasYTecnicas();
        Conclusiones();
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

        // New Styles
        style = document.Styles.AddStyle("Portada", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 11;
        style.ParagraphFormat.SpaceBefore = "0.3cm";
        style.ParagraphFormat.SpaceAfter = "0.3cm";
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Footer", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Font.Color = Colors.Gray;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
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


        // Pie de pagina

        var table = section.AddTable();
        table.Format.SpaceBefore = "5cm";
        table.AddColumn(120);
        table.AddColumn(160);
        table.AddColumn(120);
        var row = table.AddRow();
        var parrFooter = row.Cells[0].AddParagraph();
        var img = parrFooter.AddImage(pathImage + "/logo.png");
        img.Width = "3cm";
        parrFooter = row.Cells[1].AddParagraph("Certificación de Ascensores S.A.\nCalle Tabancura N° 1613 Dpto. 701 Block C. Vitacura – Santiago\nTelf. (+56) 232273961 Cel. (+56) 944821821\nEmail: contacto@certasc.cl\nwww.certasc.cl");
        parrFooter.Style = "Footer";
        var t = row.Cells[2].Elements.AddTable();
        t.Borders.Visible = true;
        t.Borders.Color = Colors.Gray;
        t.Borders.Width = 2;
        t.AddColumn(50);
        t.AddColumn(50);
        var r = t.AddRow();
        r.Cells[0].AddParagraph("VERSIÓN");
        r.Cells[1].AddParagraph("1.0");
        r = t.AddRow();
        r.Cells[0].AddParagraph("Fecha Aprobación");
        r.Cells[1].AddParagraph("01-03-2017");
        r = t.AddRow();
        r.Cells[0].AddParagraph("Código");
        r.Cells[1].AddParagraph("F-11");
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
        parr = section.AddParagraph(string.Format("Cliente: {0}", Inspeccion.NombreEdificio));
        parr = section.AddParagraph("Presente");
        parr = section.AddParagraph(string.Format("Estimados Señores: {0}", Inspeccion.Servicio.Cliente.Nombre));
        parr = section.AddParagraph(string.Format("De acuerdo con la inspección realizada el día {0} en el {1}, se envía informe {2}, con el resultado de la revisión técnica y normativa, detallando las no conformidades que deberán ser regularizadas para inicial el proceso de certificación.", 
            Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"),
            Inspeccion.NombreEdificio,
            Inspeccion.IT));
        parr = section.AddParagraph("Quedamos a su disposición y atentos a cualquier consulta.");
        parr = section.AddParagraph("Saluda atentamente,\nDpto.Técnico Ingeniería CertAsc S.A.");
        
        Paragraph tableTitle = section.AddParagraph(string.Format("INFORME DE AUDITORÍA TÉCNICA E INSPECCIÓN DEL {0}", Inspeccion.Aparato.Nombre.ToUpper()));
        tableTitle.Style = "Heading2";
        Table table1 = section.AddTable();
        table1.Borders.Visible = true;
        table1.Borders.Color = Colors.LightGray;
        table1.AddColumn(200);
        table1.AddColumn(90);
        table1.AddColumn(200);
    
        Row row = table1.AddRow();
        row.Format.Font.Bold = true;
        row.Format.Alignment = ParagraphAlignment.Center;
        row.VerticalAlignment = VerticalAlignment.Center;
        row.TopPadding = 5;
        row.BottomPadding = 5;
        row.Cells[0].MergeRight = 2;
        Paragraph parrafo1 = row.Cells[0].AddParagraph(string.Format("CONTROL DE GESTIÓN"));
        row.Cells[0].Shading.Color = Colors.LightGray;
        row = table1.AddRow();
        row.Cells[0].AddParagraph("Fecha de emisión");
        row.Cells[1].AddParagraph(Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"));
        row.Cells[2].AddParagraph(string.Format("Emitido por: {0} {1}", Inspeccion.Usuario.Nombre, Inspeccion.Usuario.Apellido));

        row = table1.AddRow();
        row.Cells[0].AddParagraph("Fecha de revisión");
        row.Cells[1].AddParagraph(Inspeccion.FechaRevision.Value.ToString("dd-MM-yyyy"));
        row.Cells[2].AddParagraph(string.Format("Revisado por: {0} {1}", Inspeccion.Usuario1.Nombre, Inspeccion.Usuario1.Apellido));

        row = table1.AddRow();
        row.Cells[0].AddParagraph("Fecha de aprobación");
        row.Cells[1].AddParagraph(Inspeccion.FechaRevision.Value.ToString("dd-MM-yyyy"));
        row.Cells[2].AddParagraph(string.Format("Aprobado por por: {0} {1}", Inspeccion.Usuario11.Nombre, Inspeccion.Usuario11.Apellido));

        row = table1.AddRow();
        row.Cells[0].AddParagraph("Fecha de entrega");
        row.Cells[1].AddParagraph(Inspeccion.FechaEntrega.Value.ToString("dd-MM-yyyy"));
        row.Cells[2].AddParagraph(string.Format("Cliente por: {0}", Inspeccion.Destinatario));

        var table2 = section.AddTable();
        table2.AddColumn(200);
        table2.AddColumn(290);
        row = table2.AddRow();
        row.Cells[0].AddParagraph("DATOS BÁSICOS");
        row.Cells[0].Shading.Color = Colors.LightGray;
        row.Cells[0].MergeRight = 1;

        row = table2.AddRow();
        parr = row.Cells[0].AddParagraph("Nº Informe");
        parr = row.Cells[1].AddParagraph(Inspeccion.IT);

        row = table2.AddRow();
        parr = row.Cells[0].AddParagraph("Dirección");
        parr = row.Cells[1].AddParagraph(Inspeccion.Ubicacion);

        row = table2.AddRow();
        parr = row.Cells[0].AddParagraph("Nombre Inspector");
        parr = row.Cells[1].AddParagraph(Inspeccion.Usuario.Nombre + " " + Inspeccion.Usuario.Apellido);

        row = table2.AddRow();
        parr = row.Cells[0].AddParagraph("Fecha de inspección");
        parr = row.Cells[1].AddParagraph(Inspeccion.FechaInspeccion.Value.ToString("dd-MM-yyyy"));

        row = table2.AddRow();
        parr = row.Cells[0].AddParagraph("Etapa del proceso"); // ?
        parr = row.Cells[1].AddParagraph(string.Empty);

        row = table2.AddRow();
        var normas = Inspeccion.InspeccionNorma.Select(s => s.Norma.Nombre).ToArray();

        parr = row.Cells[0].AddParagraph("Norma aplicada");
        parr = row.Cells[1].AddParagraph(string.Join("; ", normas));

        var table3 = section.AddTable();
        table3.AddColumn(200);
        table3.AddColumn(290);

        row = table3.AddRow();
        row.Cells[0].AddParagraph("DATOS DEL ASCENSOR");
        row.Cells[0].Shading.Color = Colors.LightGray;
        row.Cells[0].MergeRight = 1;

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Nombre edificio");
        parr = row.Cells[1].AddParagraph(Inspeccion.NombreEdificio);

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Número único del elevador");
        parr = row.Cells[1].AddParagraph(Inspeccion.Numero);

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Equipo Nº");
        parr = row.Cells[1].AddParagraph(Inspeccion.Numero); // ?

        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Destino de uso del elevador");
        parr = row.Cells[1].AddParagraph(Inspeccion.DestinoProyecto.Descripcion);


        row = table3.AddRow();
        parr = row.Cells[0].AddParagraph("Altura en pisos");
        parr = row.Cells[1].AddParagraph(Inspeccion.AlturaPisos == null ? string.Empty : Inspeccion.AlturaPisos.ToString());


        // Tabla 3

        var insp = Inspeccion.Fase == 1 ? Inspeccion : Inspeccion.Inspeccion2;




        // Especificos Tabla 3
        var especificosT2 = insp.ValoresEspecificos.Where(w => w.Especificos.NroTable == 2).OrderBy(o => o.EspecificoID);
        foreach (var e in especificosT2)
        {
            row = table3.AddRow();
            row.Cells[0].AddParagraph(e.Especificos.Nombre);
            row.Cells[1].AddParagraph(e.Valor);
            row.TopPadding = 2;
            row.BottomPadding = 2;
        }

        Table table4 = section.AddTable();
        table4.Borders.Visible = true;
        table4.KeepTogether = true;
        table4.Borders.Color = Colors.LightGray;
        table4.AddColumn(245);
        table4.AddColumn(245);
       
        Row row2 = table4.AddRow();
        row2.Format.Font.Bold = true;
        row2.Format.Alignment = ParagraphAlignment.Center;
        row2.VerticalAlignment = VerticalAlignment.Center;
        row2.TopPadding = 5;
        row2.BottomPadding = 5;
        row2.Cells[0].MergeRight = 1;

        parrafo1 = row2.Cells[0].AddParagraph("CARACTERÍSTICAS GENERALES");
        row2.Cells[0].Shading.Color = Colors.LightGray;
        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Nombre del Proyecto");
        row2.Cells[1].AddParagraph(insp.NombreProyecto ?? string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Destino del Proyecto");
        row2.Cells[1].AddParagraph(insp.DestinoProyectoID == null ? string.Empty : insp.DestinoProyecto.Descripcion);
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

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Fecha de inicio del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph("En proceso de certificación");
        row2.TopPadding = 2;
        row2.BottomPadding = 2;

        row2 = table2.AddRow();
        row2.Cells[0].AddParagraph("Fecha de vencimiento del Certificado de Inspección de Certificación");
        row2.Cells[1].AddParagraph(insp.FechaVencimientoCertificado.HasValue ? insp.FechaVencimientoCertificado.Value.ToString("dd-MM-yyyy") ?? string.Empty : string.Empty);
        row2.TopPadding = 2;
        row2.BottomPadding = 2;
    }

    public void TerminosYDefiniciones()
    {
        Section section = document.AddSection();
        Paragraph title = section.AddParagraph("TÉRMINOS Y DEFINICIONES");
        title.Style = "Heading1";
        var table = section.AddTable();
        table.AddColumn(200);
        table.AddColumn(290);
        
        using (var db = new CertelEntities())
        {
            var terminos = db.TerminosYDefiniciones
                            .Where(w => w.NormaID == NormaPrincipal)
                            .ToList();
            foreach (var t in terminos)
            {
                var row = table.AddRow();
                row.Cells[0].AddParagraph(t.Termino);
                row.Cells[1].AddParagraph(t.Definicion.TrimEnd());
            }
        }

        title = section.AddParagraph("CRITERIOS DE CALIFICACIÓN DE DEFECTOS SEGÚN D.S. N°08 (V y U)");

        var parr = section.AddParagraph("D.S. N°08 (V y U) de fecha 28 de agosto de 2017, Modifica Decreto Supremo N° 47, de Vivienda y Urbanismo, de 1992, Ordenanza General de urbanismo y Construcciones en materia de Ascensores.");

        parr = section.AddParagraph("Artículo 4°, Para los efectos de lo dispuesto en el numeral 4, del párrafo décimo cuarto del artículo 5.9.5. de la Ordenanza General de Urbanismo y Construcciones se actualizan para calificar los defectos encontrados en las instalaciones al momento de efectuar la inspección que antecede a la certificación, estos serán calificados como defectos graves y defectos leves.");

        parr = section.AddParagraph("DEFECTO GRAVE: Es todo aquél que constituye un riesgo para la seguridad de las personas, del personal técnico que mantiene las respectivas instalaciones, o de la instalación propiamente tal.");

        parr = section.AddParagraph("En virtud de lo anterior, será considerado como grave todo aquel defecto que altere o pueda alterar el correcto funcionamiento de cualquiera de los sistemas o componentes de la respectiva instalación, señalados a continuación, cuando pueda causar un accidente por cizallamiento, aplastamiento, caída, choque, atrapamiento, fuego o choque eléctrico: ");

        parr = section.AddParagraph("• Sistema de apertura de puertas, contactos de seguridad y dispositivos de enclavamiento.\n• Conjunto limitador de velocidad y paracaídas del equipo.\n• Sistemas de frenos del equipo.\n• Sistemas de suspensión y polea motriz, en especial cuando estos no cumplan con las disposiciones de seguridad especificadas por el fabricante.\n• Línea eléctrica o circuito de seguridad, incluidos los dispositivos de final de recorrido.\n• Registros carpeta de ascensores.");

        parr = section.AddParagraph("DEFECTO LEVE: Es todo aquel no calificable como grave, y que por sí solo no significa un riesgo para la seguridad de las personas, para el personal técnico que mantiene las respectivas instalaciones, o para la instalación propiamente tal.");

        parr = section.AddParagraph("En caso de que, conforme a las normas técnicas oficiales vigentes aplicables a la respectiva instalación, haya razones técnicas por las cuales estos defectos no puedan ser subsanados, el certificador deberá determinar fundadamente una solución alternativa para cada defecto, de carácter permanente, así como el plazo de ejecución de la misma solución, lo que deberá quedar detallado en un informe de defectos leves que se adjuntará a la certificación.");
    }
    public void ResultadosInspeccion()
    {
        Section section = document.AddSection();
        
        using (var db = new CertelEntities())
        {
            var parr = section.AddParagraph("TERMINOLOGÍA: Defecto grave (DG), Defecto leve, (DL), No aplica (N/A), Cumple con el requisito (OK).");

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
                table.Borders.Color = Colors.LightGray;
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
                
                row.Cells[0].AddParagraph(t.Texto);

                // ENCABEZADO
                row = table.AddRow();
                row.Cells[0].MergeDown = 1;
                row.Cells[0].AddParagraph(string.Format("{0}", n.Nombre));
                row.Cells[1].MergeDown = 1;
                row.Cells[1].AddParagraph("Criterio de aceptación");
                row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
                row.Cells[2].AddParagraph("Observaciones");


                // SUB-ENCABEZADO
                row = table.AddRow();
                if (Inspeccion.Fase == 1)
                {
                    row.Cells[2].AddParagraph("OK");
                    row.Cells[3].AddParagraph("DG");
                    row.Cells[4].AddParagraph("DL");
                    row.Cells[5].AddParagraph("N/A");
                }
                else
                {
                    row.Cells[2].AddParagraph("Corregido");
                    row.Cells[3].AddParagraph("No Corregido");
                }
                



                
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
                        var parr1 = cRow.Cells[1].AddParagraph(c.Descripcion);
                        parr1.Style = "Caract";
                        cRow.Cells[0].AddParagraph(string.Format("{0}", r.Descripcion));
                        cRow.Cells[0].MergeDown = carsCount - 1;

                        var cumplimiento = c.Cumplimiento
                                            .Where(w => Inspeccion.Fase == 1
                                                    ? w.InspeccionID == Inspeccion.ID
                                                    : w.InspeccionID == Inspeccion.InspeccionFase1)
                                            .FirstOrDefault();
                        if (cumplimiento == null)
                            continue;

                        var index = 0;
                        switch (cumplimiento.EvaluacionID) {
                            case 1: index = 2; break;
                            case 2: index = 4; break;
                            case 3: index = 3; break;
                            case 4: index = 5; break;
                            case 5: index = 2; break;
                            case 6: index = 3; break;
                        }
                        parr1 = cRow.Cells[index].AddParagraph(cumplimiento == null ? string.Empty : "X");
                        parr1.Style = "Parrafo";
                        parr1.Format.Alignment = ParagraphAlignment.Center;
                        parr1.Style = "Carac";
                        parr1.Format.Alignment = ParagraphAlignment.Justify;
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
                                    parr1 = cRow.Cells[3].AddParagraph("X");
                                    
                                }

                            }

                        }


                    }

                }

            }
            var normasAsociadas = n.NormasAsociadas;

            foreach (var nor in normasAsociadas)
            {
                var na = db.Norma.Find(nor.NormaSecundariaID);
                var titles = na.Titulo.ToList();
                foreach (var t in titles)
                {

                    Table table = section.AddTable();
                    table.Borders.Visible = true;
                    table.Borders.Color = Colors.LightGray;
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
                    row.Cells[0].AddParagraph(string.Format("{0}", na.TituloRegulacion));

                    row = table.AddRow();
                    row.Cells[0].AddParagraph(na.Nombre);
                    row.Cells[1].AddParagraph("Criterio de aceptación");
                    row.Cells[2].AddParagraph("Observaciones");
                    row.Cells[3].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;


                    row = table.AddRow();
                    row.Format.Font.Bold = true;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = VerticalAlignment.Center;
                    if (Inspeccion.Fase == 1)
                    {
                        row.Cells[2].AddParagraph("OK");
                        row.Cells[3].AddParagraph("DG");
                        row.Cells[4].AddParagraph("DL");
                        row.Cells[5].AddParagraph("N/A");
                    }
                    else
                    {
                        row.Cells[2].AddParagraph("Corregido");
                        row.Cells[3].AddParagraph("No Corregido");
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
                                case 1: index = 2; break;
                                case 2: index = 4; break;
                                case 3: index = 3; break;
                                case 4: index = 5; break;
                                case 5: index = 2; break;
                                case 6: index = 3; break;
                            }
                            parr1 = cRow.Cells[index].AddParagraph(cumplimiento == null ? string.Empty : "X");
                            parr1.Style = "Parrafo";
                            parr1.Format.Alignment = ParagraphAlignment.Center;
                            if (cumplimiento.EvaluacionID == 3)
                            {
                                parr1.Format.Font.Color = Colors.Blue;
                                if (Inspeccion.Fase > 1)
                                {
                                    var corregido = c.Cumplimiento
                                                        .Where(w => w.InspeccionID == Inspeccion.ID)
                                                        .FirstOrDefault();
                                    parr1 = cRow.Cells[corregido != null ? 3 : 4].AddParagraph("X");
                                    parr1.Style = "Carac";
                                }

                            }
                        }
                    }
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
        
        Paragraph texto = section.AddParagraph(string.Format("Las siguientes observaciones deben ser corregidas para que el elevador quede en norma, y pueda ser certificado:", Inspeccion.Aparato.Nombre));
        texto.Style = "Parrafo";
        title = section.AddParagraph(string.Format("{0}.{1} OBSERVACIONES POR NORMA", point, subpoint));
        title.Style = "Heading2";
        
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
                texto = section.AddParagraph(string.Format("La OGUC (Ordenanza General de Urbanismo y Construcciones) en el Artículo 5.1.6, Numeral 13, indica que los elevadores deben disponer de una carpeta cero (o carpeta del elevador), este requisito es reafirmado por el punto Registros, de la norma {0} que indica la documentación necesaria que debe disponer dicha carpeta.", NormaPrincipalNombre));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Es importante que tanto la administración del Edificio {0}, como la empresa mantenedora, colaboren en la implementación de la carpeta cero,  ya que existen en ella documentos que servirán para inscribir el ascensor en la DOM (Dirección de Obras Municipales) según la indicación de la OGUC Artículo 5.9.5. Numeral 1, mediante una identificación con número único de registro del elevador.", Inspeccion.NombreEdificio));
                texto.Style = "Parrafo";
                if(Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("El {0} N° {1}, en su estado actual, NO CALIFICA PARA LA CERTIFICACIÓN POR DEFECTOS GRAVES, según  las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas. Se recomienda  corregir las no conformidades y observaciones técnicas según la norma {2} señaladas en los puntos 4 y 5 del presente informe para que el {0} pueda cumplir con las normas Chilenas y pueda certificarse sin observaciones.", Inspeccion.Aparato.Nombre, Inspeccion.Numero, NormaPrincipalNombre));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Se da un plazo de {0} días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de no conformidades del {1}.", Inspeccion.DiasPlazo == null ? "90" : Inspeccion.DiasPlazo.ToString(), Inspeccion.Aparato.Nombre));
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
                texto = section.AddParagraph(string.Format("Según la evaluación del (Ascensor o Montacargas), se encuentran N° de hallazgos denominados “Defectos Leves” (DL) correspondiente al N°% y de N°s de “Conformidades” (OK)correspondiente a N°% de un total de N° de requisitos aplicados normativamente."));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("El Ascensor o Montacargas N°------------------------------- en su estado actual, califica para la certificación con defectos leves, según las disposiciones contenidas en la Ley 20.296 y el D.S.N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S.N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas. "));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("Se deben corregir los defectos leves y observaciones técnicas registradas en este informe según exigencias de la normas ------------------------------ señaladas en el presente informe para que el(ascensor o montacargas) pueda cumplir con las normas chilenas y certificarse sin observaciones."));
                texto.Style = "Parrafo";
                
                texto.Style = "Parrafo";
                if (Inspeccion.Fase == 1)
                {
                    texto = section.AddParagraph(string.Format("Se otorga un plazo de 90 días corridos a partir de la fecha del envío de este informe para realizar trabajos correspondientes a las mejoras y/o levantamiento de los hallazgos encontrados."));
                    texto.Style = "Parrafo";
                    texto = section.AddParagraph(string.Format("Cumplido este plazo, se programará en conjunto con el cliente la Etapa II del servicio, para revisar si lo solicitado/sugerido en este informe, fue realizado, y así verificar si el equipo califica o no para su certificación sin observaciones."));
                    texto.Style = "Parrafo";


                    // HASTA AQUI LLEGUE
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
                texto = section.AddParagraph(string.Format("En conformidad a las disposiciones contenidas en la Ley 20.296 y el D.S. N° 47 “Ordenanza General de Urbanismo y Construcciones” OGUC, modificado por el D.S. N° 37 – D.O. 22.03.2016 y en cumplimiento del Artículo 5.9.5 numeral 4: Certificación de ascensores, montacargas y escaleras o rampas mecánicas, se acredita mediante inspección técnica y normativa, que la instalación del Ascensor - Montacargas, cumple con los requisitos de instalación y de las seguridades en conformidad con las normas (Normas seleccionadas en la inspección) aplicadas. Por lo tanto, se acredita que el elevador ha sido adecuadamente mantenido y que se encuentran en condiciones de seguir funcionando."));
                texto.Style = "Parrafo";
                texto = section.AddParagraph(string.Format("El(ascensor o montacargas) N° ----------------------------------------------, califica para la certificación, cumpliendo con la Ley 20.296. El certificado de inspección técnica y normativa denominado Certificado de Inspección Electromecánico, deberá ser ingresado a la Dirección de Obras Municipales respectiva, por el propietario o por el administrador, según corresponda, antes del vencimiento del plazo que tiene la instalación para certificarse, y dentro de un plazo no superior a 30 días contados desde la fecha de emisión de la certificación. Se procederá entonces, a emitir el certificado de inspección electromecánico y de experiencia del elevador, el que estará disponible para su despacho en un plazo máximo de 5 días hábiles."));
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
        var section = document.AddSection();
        var table = section.AddTable();
        table.Borders.Visible = true;
        table.Borders.Width = 2;
        table.Borders.Color = Colors.Gray;
        if (Inspeccion.Fase == 1)
        {
            table.AddColumn(160);
            table.AddColumn(160);
            table.AddColumn(40);
            table.AddColumn(40);
            table.AddColumn(40);
            table.AddColumn(40);
        }
        else
        {
            table.AddColumn(160);
            table.AddColumn(160);
            table.AddColumn(80);
            table.AddColumn(80);
        }
        

        var row = table.AddRow();
        row.Cells[0].AddParagraph("RESUMEN DE CUMPLIMIENTOS");
        row.Cells[0].MergeRight = 5;


        row = table.AddRow();
        row.Cells[0].AddParagraph("TOTAL REQUISITOS NORMATIVOS");
        row.Cells[1].AddParagraph("?"); // ?
        row.Cells[2].AddParagraph("CUMPLIMIENTOS");
        row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;


        row = table.AddRow();
        row.Cells[0].AddParagraph("TOTAL REQUISITOS APLICADOS");
        row.Cells[1].AddParagraph("?"); // ?
 

        if (Inspeccion.Fase == 1)
        {
            row.Cells[2].AddParagraph("OK");
            row.Cells[3].AddParagraph("DG");
            row.Cells[4].AddParagraph("DL");
            row.Cells[5].AddParagraph("N/A");
        }
        else
        {
            row.Cells[2].AddParagraph("Corregido");
            row.Cells[3].AddParagraph("No Corregido");

        }

        row = table.AddRow();
        row.Cells[0].AddParagraph("CANTIDAD POR CALIFICACIÓN");
        row.Cells[1].AddParagraph(string.Empty);

        if (Inspeccion.Fase == 1) // FIXME
        {
            row.Cells[2].AddParagraph("1");
            row.Cells[3].AddParagraph("1");
            row.Cells[4].AddParagraph("1");
            row.Cells[5].AddParagraph("1");
        }
        else
        {
            row.Cells[2].AddParagraph("1");
            row.Cells[3].AddParagraph("1");

        }

        row = table.AddRow();
        row.Cells[0].AddParagraph("PORCENTAJE POR CALIFICACIÓN");
        row.Cells[1].AddParagraph(string.Empty);

        if (Inspeccion.Fase == 1) // FIXME
        {
            row.Cells[2].AddParagraph("33%");
            row.Cells[3].AddParagraph("33%");
            row.Cells[4].AddParagraph("33%");
            row.Cells[5].AddParagraph("33%");
        }
        else
        {
            row.Cells[2].AddParagraph("33%");
            row.Cells[3].AddParagraph("33%");

        }

        // fila vacia
        row = table.AddRow();
        row.Cells[0].AddParagraph(string.Empty);
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;

        row = table.AddRow();
        row.Cells[0].AddParagraph("RESUMEN DE OBSERVACIONES NORMATIVAS Y TÉCNICAS");
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;

        row = table.AddRow();
        row.Cells[0].AddParagraph("Las siguientes observaciones deben ser corregidas para que el elevador quede en norma, y pueda ser certificado:");
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;

        row = table.AddRow();
        row.Cells[0].AddParagraph("OBSERVACIONES POR NORMA");
        row.Cells[0].MergeRight = Inspeccion.Fase == 1 ? 5 : 3;


        row = table.AddRow();
        var parr = row.Cells[0].AddParagraph("REQUISITO");
        parr = row.Cells[1].AddParagraph("DESCRIPCIÓN");
        parr = row.Cells[2].AddParagraph("IMAGEN");
        row.Cells[2].MergeRight = Inspeccion.Fase == 1 ? 3 : 1;
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
                row = table.AddRow();
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                parr = row.Cells[0].AddParagraph(puntoNC);
                parr = row.Cells[1].AddParagraph(nc.Observacion ?? string.Empty);
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
            section.AddPageBreak();
        }
        if (noCumplimientoConFoto.Count() > 0)
        {
            foreach (var nc in noCumplimientoConFoto)
            {
                row = table.AddRow();
                if (count == 2)
                {
                    section.AddPageBreak();
                    count = 0;
                }
                var puntoNC = nc.Requisito.Replace("\n", " ").TrimEnd();
                parr = row.Cells[0].AddParagraph(puntoNC);
                parr = row.Cells[1].AddParagraph(nc.Observacion ?? string.Empty);
                parr = row.Cells[2].AddParagraph();
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
                    image.Width = "4cm";
                    numberfoto++;

                }
                count++;
            }
        }

        //var observacionesTecnicas = insp.ObservacionTecnica;
        //if (observacionesTecnicas.Count == 0)
        //    return;

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