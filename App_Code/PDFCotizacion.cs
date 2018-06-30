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
using System.Web.Script.Serialization;
using System.Threading;
using System.Globalization;
/// <summary>
/// Descripción breve de PDFCotizacion
/// </summary>
public class PDFCotizacion
{
    public Cotizacion Cotizacion { get; set; }
    public Document document { get; set; }
    public Section section { get; set; }
    public Table TableContainer { get; set; }
    public Column ContainerColumn { get; set; }
    public string FileName { get; set; }
    public ConfigVariables config = new ConfigVariables();
    private string format = "{0:n0}";
    public PDFCotizacion(Cotizacion cotizacion)
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
        config.SetData();
        Cotizacion = cotizacion;
        if (cotizacion.MonedaId != 1)
        {
            format = "{0:n}";
        }
        Create();
        DefineStyles();
        CreateHeader();
       // CreateIT();
        CreatePrincipal();
        var items = Cotizacion.ItemCotizacion.Count;
        var limit = items > 10 ? 11 : items + 3;
        var iteraciones = Math.Ceiling(items / (double)limit);
        for (var i = 1; i <= iteraciones; i++) {
            CreateDetalle(i, iteraciones, limit);
        }
        
        CreateFoot();
        FileName = Rendering();
        
    }
    public void Create() {
        document = new Document();
        document.Info.Title = "Cotización";
        document.DefaultPageSetup.TopMargin = "2cm";
        document.DefaultPageSetup.LeftMargin = "2cm";
        document.DefaultPageSetup.RightMargin = "2cm";
        section = document.AddSection();
        CreateTable();

    }
    public void CreateTable() {
        TableContainer = section.AddTable();
        ContainerColumn = TableContainer.AddColumn(500);
        TableContainer.Borders.Visible = true;
        TableContainer.Borders.Color = Colors.Black;
        TableContainer.Borders.Width = Unit.FromMillimeter(0.4);
        TableContainer.Borders.Distance = 50;
    }
    public void DefineStyles()
    {
        Style style = document.Styles.AddStyle("Head", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 9;
        style.ParagraphFormat.Font.Color = Colors.Gray;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Head-small", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 7;
        style.ParagraphFormat.Font.Color = Colors.Black;

        style = document.Styles.AddStyle("Title", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Shading.Color = Colors.LightGray;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Principal", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.SpaceAfter = Unit.FromMillimeter(1);
        style.ParagraphFormat.SpaceBefore = Unit.FromMillimeter(1);

        style = document.Styles.AddStyle("PrincipalDato", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = Unit.FromMillimeter(1);
        style.ParagraphFormat.SpaceBefore = Unit.FromMillimeter(1);

        style = document.Styles.AddStyle("EncabezadoDetalle", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("Detalle", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;

        style = document.Styles.AddStyle("DetalleCenter", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("DetalleRigth", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 7;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        style = document.Styles.AddStyle("DetalleRigthTotal", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 7;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        style.ParagraphFormat.SpaceAfter = 5;
        style.ParagraphFormat.SpaceBefore = 5;
        style.ParagraphFormat.Font.Bold = true;

        style = document.Styles.AddStyle("Nota", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 7;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.Font.Bold = true;

        style = document.Styles.AddStyle("ParrafoFooter", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 8;
        style.ParagraphFormat.Font.Color = Colors.DarkBlue;
        style.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("DetalleFoot", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 6;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        style.ParagraphFormat.SpaceAfter = 5;
        style.ParagraphFormat.SpaceBefore = 5;


        style = document.Styles.AddStyle("DatosTransfer", "Normal");
        style.Font.Name = "Arial";
        style.ParagraphFormat.Font.Size = 7;
        style.ParagraphFormat.Font.Color = Colors.Black;
        style.ParagraphFormat.Font.Italic = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Left;

    }
    public void CreateHeader ()
    {
        // ROWS
        Row containerRow = TableContainer.AddRow();
        Table tableInside = containerRow.Cells[0].Elements.AddTable();
        tableInside.AddColumn(500);
        string pathImage = HttpContext.Current.Server.MapPath("~/css/images/");
        var row = tableInside.AddRow();
        var parrImg = row.Cells[0].AddParagraph();
        parrImg.Format.Alignment = ParagraphAlignment.Center;
        Image image = parrImg.AddImage(pathImage + "/logo.png");
        image.Height = 100;
        
        row = tableInside.AddRow();
        var parr = row.Cells[0].AddParagraph("Rut: " + config.rut);
        parr.Style = "Head";
        row = tableInside.AddRow();
        parr = row.Cells[0].AddParagraph("Dirección: " +config.direccion);
        parr.Style = "Head";
        row = tableInside.AddRow();
        parr = row.Cells[0].AddParagraph("Giro: " + config.giro);
        parr.Style = "Head";
        row = tableInside.AddRow();
        parr = row.Cells[0].AddParagraph("Teléfono: " + config.telefono);
        parr.Style = "Head";
        row = tableInside.AddRow();
        parr = row.Cells[0].AddParagraph("Cel: " + config.cel);
        parr.Style = "Head";
        
    }
    public void CreateIT() 
    {
        Row containerRow = TableContainer.AddRow();
        var table = containerRow.Cells[0].Elements.AddTable();
        table.AddColumn(500);
        var row = table.AddRow();
        var parr = row.Cells[0].AddParagraph(string.Format("COTIZACIÓN N° {0}", Cotizacion.IT));
        parr.Style = "Title";
    }
    public void CreatePrincipal ()
    {
        Row containerRow = TableContainer.AddRow();
        var parr = containerRow.Cells[0].AddParagraph("DATOS CLIENTE");
        parr.Style = "Title";

        containerRow = TableContainer.AddRow();
        var table = containerRow.Cells[0].Elements.AddTable();
        table.AddColumn(90);
        table.AddColumn(5);
        table.AddColumn(180);
        table.AddColumn(80);
        table.AddColumn(5);
        table.AddColumn(135);

        //1
        var row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Nombre");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(Cotizacion.Cliente.Nombre.ToUpper());
        parr.Style = "PrincipalDato";
        
        //2
        parr = row.Cells[3].AddParagraph("Atención");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(Cotizacion.Cliente.NombreContacto.ToUpper());
        parr.Style = "PrincipalDato";
        
        //3
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Referencia");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(config.referencia.ToUpper());
        parr.Style = "PrincipalDato";
        
        //4
        parr = row.Cells[3].AddParagraph("Rut");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(Cotizacion.Cliente.Rut.ToUpper());
        parr.Style = "PrincipalDato";
        

        //5
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Teléfono");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(Cotizacion.Cliente.TelefonoContacto.ToUpper());
        parr.Style = "PrincipalDato";
        
        //6
        parr = row.Cells[3].AddParagraph("Email");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(Cotizacion.Cliente.EmailContacto.ToUpper());
        parr.Style = "PrincipalDato";
        

        //7
        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Fecha");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(Cotizacion.FechaCreacion.HasValue ? Cotizacion.FechaCreacion.Value.ToString("dd-MM-yyyy") : string.Empty);
        parr.Style = "PrincipalDato";
        

        containerRow = TableContainer.AddRow();
        parr = containerRow.Cells[0].AddParagraph("CERTASC");
        parr.Style = "Title";

        

        table = containerRow.Cells[0].Elements.AddTable();
        table.AddColumn(90);
        table.AddColumn(5);
        table.AddColumn(180);
        table.AddColumn(80);
        table.AddColumn(5);
        table.AddColumn(135);

        row = table.AddRow();
        var admCargo = Cotizacion.Usuario == null ? string.Empty : Cotizacion.Usuario1.Cargo;
        parr = row.Cells[0].AddParagraph(admCargo);
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        var admComercial = Cotizacion.Usuario == null ? string.Empty : (Cotizacion.Usuario1.Nombre + " " + Cotizacion.Usuario1.Apellido).ToUpper();
        parr = row.Cells[2].AddParagraph(admComercial);
        parr.Style = "PrincipalDato";

        parr = row.Cells[3].AddParagraph("Teléfono");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(Cotizacion.Usuario1.Fono ?? string.Empty);
        parr.Style = "PrincipalDato";

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Celular");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(Cotizacion.Usuario1.Celular ?? string.Empty);
        parr.Style = "PrincipalDato";

        parr = row.Cells[3].AddParagraph("Email");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        var email = Cotizacion.Usuario == null ? string.Empty : Cotizacion.Usuario1.Email;
        parr = row.Cells[5].AddParagraph(email == null ? string.Empty : email.ToUpper());
        parr.Style = "PrincipalDato";

        row = table.AddRow();
        parr = row.Cells[0].AddParagraph("Lugar de entrega");
        parr.Style = "Principal";
        parr = row.Cells[1].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[2].AddParagraph(Cotizacion.Ciudad.Descripcion.ToUpper());
        parr.Style = "PrincipalDato";

        parr = row.Cells[3].AddParagraph("Condiciones de venta");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(config.condiciones_de_venta.ToUpper());
        parr.Style = "PrincipalDato";

        row = table.AddRow();
        parr = row.Cells[3].AddParagraph("Condiciones de Pago");
        parr.Style = "Principal";
        parr = row.Cells[4].AddParagraph(":");
        parr.Style = "Principal";
        parr = row.Cells[5].AddParagraph(Cotizacion.FormaDePagoId == null ? string.Empty : Cotizacion.FormaDePago.Descripcion.ToUpper());
        parr.Style = "PrincipalDato";
    }
    public void CreateDetalle (int page, double total, int limit)
    {
        if(page > 1) {
            section.AddPageBreak();
            CreateTable();
        }
        else {
            CreateIT();
        }
        Row containerRow = TableContainer.AddRow();
        var table = containerRow.Cells[0].Elements.AddTable();
        table.AddColumn(20);
        table.AddColumn(271);
        table.AddColumn(30);
        table.AddColumn(20);
        
        table.AddColumn(45);
        table.AddColumn(45);
        table.AddColumn(65);
        table.Borders.Visible = true;
        table.Borders.Width = 0.1;
        table.Format.SpaceBefore = 10;
        var row = table.AddRow();
        var parr = row.Cells[0].AddParagraph("Item");
        parr.Style = "EncabezadoDetalle";
        row.Borders.Top.Visible = false;
        parr = row.Cells[1].AddParagraph("Detalle");
        parr.Style = "EncabezadoDetalle";
        row.Borders.Left.Visible = false;
        parr = row.Cells[2].AddParagraph("Cantidad");
        parr.Style = "EncabezadoDetalle";

        parr = row.Cells[3].AddParagraph("U/M");
        parr.Style = "EncabezadoDetalle";

        parr = row.Cells[4].AddParagraph("Valor Unitario.");
        parr.Style = "EncabezadoDetalle";

        parr = row.Cells[5].AddParagraph("Total");
        parr.Style = "EncabezadoDetalle";

        parr = row.Cells[6].AddParagraph("Entrega");
        parr.Style = "EncabezadoDetalle";

        
        
        var index = (page - 1) * limit;
        var maxHeightItem = 370;
        var hasNota = !string.IsNullOrEmpty(Cotizacion.Nota);
        var items = Cotizacion.ItemCotizacion.Skip((page - 1) * limit).Take(limit);
        var itemsLength = items.Count() + (Cotizacion.GastoOperacional.Any() && page == total ? 1 : 0) + (hasNota && page == total ? 3 : 0);
        var height = page == total ? (maxHeightItem / (itemsLength + 2)) : (maxHeightItem / itemsLength);
        //row = table.AddRow();
        //row.HeightRule = RowHeightRule.Exactly;
        //row.Height = 30;
        //row.TopPadding = 10;
        //row.BottomPadding = 10;
        //row.Borders.Bottom.Visible = false;
        //parr = row.Cells[0].AddParagraph(string.Empty);
        //parr.Style = "DetalleRigth";
        //parr = row.Cells[1].AddParagraph(string.Empty);
        //parr.Style = "DetalleRigth";
        //parr = row.Cells[2].AddParagraph(string.Empty);
        //parr.Style = "Detalle";
        //parr.Format.Alignment = ParagraphAlignment.Center;
        //parr = row.Cells[3].AddParagraph("SERVICIO DE INSPECCIÓN, REVISIÓN DE LA INSTALACIÓN Y FUNCIONAMIENTO PARA:");
        //parr.Style = "Detalle";
        //parr = row.Cells[4].AddParagraph(string.Empty);
        //parr.Style = "DetalleCenter";
        //parr = row.Cells[5].AddParagraph(string.Empty);
        //parr.Style = "DetalleRigth";
        //parr = row.Cells[6].AddParagraph(string.Empty);
        //parr.Style = "DetalleRigth";
        //row.Cells[0].Borders.Left.Visible = false;
        //row.Cells[6].Borders.Right.Visible = false;
        foreach (var item in items)
        {
            index++; 

            row = table.AddRow();
            row.HeightRule = RowHeightRule.Exactly;
            row.Height = height;
            row.TopPadding = 10;
            row.BottomPadding = 10;
            row.Borders.Bottom.Visible = true;
            parr = row.Cells[0].AddParagraph(index.ToString());
            parr.Style = "DetalleRigth";
            parr = row.Cells[1].AddParagraph(item.DescripcionEditada);
            parr.Style = "Detalle";
            parr = row.Cells[2].AddParagraph(item.Cantidad.ToString());
            parr.Style = "DetalleRigth";
            parr = row.Cells[3].AddParagraph(config.unidad_medida);
            parr.Style = "Detalle";
            parr.Format.Alignment = ParagraphAlignment.Center;

            parr = row.Cells[4].AddParagraph(string.Format(format, item.Unitario));
            parr.Style = "DetalleRigth";

            parr = row.Cells[5].AddParagraph(string.Format(format, (item.Unitario * item.Cantidad)));
            parr.Style = "DetalleRigth";

            parr = row.Cells[6].AddParagraph(index == 1 ? config.fecha_entrega.ToUpper() : string.Empty);
            parr.Style = "DetalleCenter";
            
            
            row.Cells[0].Borders.Left.Visible = false;
            row.Cells[6].Borders.Right.Visible = false;
        }

        if (page < total)
        {
            return;
        }
        // COSTOS LOGISTICOS
        if(Cotizacion.GastoOperacional.Any())
        {
            index++;
            row = table.AddRow();
            //row.HeightRule = RowHeightRule.Exactly;
            row.Height = height;
            row.TopPadding = 10;
            row.BottomPadding = 10;
            row.Borders.Bottom.Visible = false;
            parr = row.Cells[0].AddParagraph(index.ToString());
            parr.Style = "DetalleRigth";
            parr = row.Cells[1].AddParagraph("COSTOS LOGÍSTICOS");
            parr.Style = "Detalle";

            parr = row.Cells[2].AddParagraph(string.Empty);
            parr.Style = "DetalleRigth";
            parr = row.Cells[3].AddParagraph(string.Empty);
            parr.Style = "Detalle";
            parr.Format.Alignment = ParagraphAlignment.Center;
            
            parr = row.Cells[4].AddParagraph(string.Empty);
            parr.Style = "Detalle";
            parr = row.Cells[5].AddParagraph(string.Format(format, Cotizacion.GastoOperacional.Sum(s => s.Valor * s.Cantidad)));
            parr.Style = "DetalleRigth";
            parr = row.Cells[6].AddParagraph(string.Empty);
            parr.Style = "DetalleRigth";
            
            row.Cells[0].Borders.Left.Visible = false;
            row.Cells[6].Borders.Right.Visible = false;
            row.Borders.Bottom.Visible = true;
            row.Borders.Bottom.Width = 0.5;
        }
        
        // NOTA:

        //if (!string.IsNullOrEmpty(Cotizacion.Nota))
        //{
        //    row = table.AddRow();
        //    //row.HeightRule = RowHeightRule.Exactly;
        //    row.Height = height * 3;
        //    row.TopPadding = 10;
        //    row.BottomPadding = 10;
        //    row.Borders.Bottom.Visible = false;
        //    row.Borders.Top.Visible = false;
        //    row.Cells[0].Borders.Left.Visible = false;
        //    row.Cells[6].Borders.Right.Visible = false;
        //    parr = row.Cells[3].AddParagraph("NOTA: " + Cotizacion.Nota.ToUpper());
        //    parr.Style = "Nota";
        //}
       
        // Foot


        row = table.AddRow();
        row.BottomPadding = 5;
        row.TopPadding = 5;
        row.Borders.Bottom.Visible = false;
        row.Borders.Left.Visible = false;
        row.Borders.Top.Visible = true;
        row.Cells[0].MergeRight = 3;
        var table2 = row.Cells[0].Elements.AddTable();
        table2.Borders.Visible = false;
        table2.AddColumn(70);
        table2.AddColumn(250);

        var row2 = table2.AddRow();
        var parr2 = row2.Cells[0].AddParagraph("Validez de la oferta:");
        parr2.Style = "DetalleFoot";
        parr2 = row2.Cells[0].AddParagraph("Moneda:");
        parr2.Style = "DetalleFoot";
        parr2 = row2.Cells[0].AddParagraph("Observaciones:");
        parr2.Style = "DetalleFoot";
        parr2 = row2.Cells[1].AddParagraph("HASTA EL " + Cotizacion.FechaValidez.Value.ToLongDateString().ToUpper());
        parr2.Style = "DetalleFoot";
        parr2 = row2.Cells[1].AddParagraph(Cotizacion.Moneda.Descripcion.ToUpper());
        parr2.Style = "DetalleFoot";
        parr2 = row2.Cells[1].AddParagraph(Cotizacion.Observacion == null ? string.Empty : Cotizacion.Observacion.ToUpper());
        parr2.Style = "DetalleFoot";

        parr = row.Cells[4].AddParagraph(string.Format(format, (Cotizacion.Valor + Cotizacion.Recargo)));
        parr.Style = "DetalleRigthTotal";
        if (Cotizacion.Descuento != null && Cotizacion.Descuento > 0)
        {
            parr = row.Cells[4].AddParagraph(string.Format(format, Cotizacion.Descuento));
            parr.Style = "DetalleRigthTotal";
        }
        parr = row.Cells[4].AddParagraph("0");
        parr.Style = "DetalleRigthTotal";
        parr = row.Cells[4].AddParagraph(string.Format(format, Cotizacion.Total));
        parr.Style = "DetalleRigthTotal";

        parr = row.Cells[6].AddParagraph("TOTAL NETO");
        parr.Style = "DetalleRigthTotal";
        if(Cotizacion.Descuento != null && Cotizacion.Descuento > 0)
        {
            parr = row.Cells[6].AddParagraph("DESCUENTO");
            parr.Style = "DetalleRigthTotal";
        }
        
        parr = row.Cells[6].AddParagraph("IVA");
        parr.Style = "DetalleRigthTotal";
        parr = row.Cells[6].AddParagraph("TOTAL EXENTO");
        parr.Style = "DetalleRigthTotal";
        row.Cells[4].MergeRight = 1;
        

        

    }
    public void CreateFoot ()
    {
        Row containerRow = TableContainer.AddRow();
        var table = containerRow.Cells[0].Elements.AddTable();
        var column = table.AddColumn(500);
        var row = table.AddRow();
        // Transferencia
        //row.HeightRule = RowHeightRule.Exactly;
        row.Height = 60;
        row.TopPadding = 3;
        row.BottomPadding = 3;
        row.Borders.Bottom.Visible = true;
        row.Borders.Top.Visible = false;
        row.Cells[0].Borders.Left.Visible = false;
        var parr = row.Cells[0].AddParagraph("DATOS TRANSFERENCIA O DEPÓSITO (NO RETIRAMOS CHEQUES)");
        parr.Style = "DatosTransfer";
        parr.Format.Font.Bold = true;
        parr.Format.Font.Italic = false;
        parr.Format.Font.Underline = Underline.Single;
        parr = row.Cells[0].AddParagraph(config.datos_transferencia_nombre);
        parr.Style = "DatosTransfer";

        parr = row.Cells[0].AddParagraph(config.datos_transferencia_banco);
        parr.Style = "DatosTransfer";

        parr = row.Cells[0].AddParagraph(config.datos_transferencia_cta);
        parr.Style = "DatosTransfer";

        parr = row.Cells[0].AddParagraph(config.datos_transferencia_rut);
        parr.Style = "DatosTransfer";

        parr = row.Cells[0].AddParagraph(config.datos_transferencia_email);
        parr.Style = "DatosTransfer";
        parr = row.Cells[0].AddParagraph("Obs: También existe la opción de pagar con tarjetas bancarias");
        parr.Style = "DatosTransfer";
        parr = row.Cells[0].AddParagraph(config.webPage);
        parr.Style = "ParrafoFooter";
        parr = row.Cells[0].AddParagraph(config.email_contacto);
        parr.Style = "ParrafoFooter";
        parr = row.Cells[0].AddParagraph(config.telefono);
        parr.Style = "ParrafoFooter";

    }
    public string Rendering()
    {
        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true)
        {
            Document = document
        };
        pdfRenderer.RenderDocument();
        string filename = string.Format("Cotizacion_{0}.pdf", Cotizacion.IT);
        string basePath = HttpContext.Current.Server.MapPath("~/cotizaciones/");
        string path = basePath + filename;
        pdfRenderer.PdfDocument.Save(path);
        return filename;
    }
}