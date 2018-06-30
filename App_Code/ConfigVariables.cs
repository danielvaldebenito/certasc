using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;

/// <summary>
/// Descripción breve de ConfigVariables
/// </summary>
public class ConfigVariables
{
    public string nombre { get; set; }
    public string nombre_largo { get; set; }
    public string rut { get; set; }
    public string direccion { get; set; }
    public string giro { get; set; }
    public string telefono { get; set; }
    public string cel { get; set; }
    public string fecha_entrada_vigencia { get; set; }
    public string referencia { get; set; }
    public string condiciones_de_venta { get; set; }
    public string condiciones_de_pago { get; set; }
    public string fecha_entrega { get; set; }
    public string unidad_medida { get; set; }
    public string mail_envio_cotizacion_subject { get; set; }
    public string mail_envio_cotizacion_text { get; set; }
    public string public_logo { get; set; }
    public string agendamiento_mail { get; set; }
    public string cotizacion_mail { get; set; }
    public string mail_seguimiento_cotizacion { get; set; }
    public string mail_post_venta { get; set; }
    public string datos_transferencia_nombre { get; set; }
    public string datos_transferencia_banco { get; set; }
    public string datos_transferencia_cta { get; set; }
    public string datos_transferencia_rut { get; set; }
    public string datos_transferencia_email { get; set; }
    public string webPage { get; set; }
    public string email_contacto { get; set; }
    public string nota_descripcion { get; set; }
    public string observacion { get; set; }
    public string mail_alerta_fase_2 { get; set; }
    public ConfigVariables()
    {
        //SetData();
    }

    public bool SetData () {
        var baseUrl = HttpContext.Current.Server.MapPath("~/");
        using (StreamReader r = new StreamReader(baseUrl + "config-variables.json"))
        {
            string json = r.ReadToEnd();
            var deserializer = new JavaScriptSerializer();
            var config = deserializer.Deserialize<ConfigVariables>(json);
            nombre = config.nombre;
            nombre_largo = config.nombre_largo;
            rut = config.rut;
            direccion = config.direccion;
            giro = config.giro;
            telefono = config.telefono;
            cel = config.cel;
            fecha_entrada_vigencia = config.fecha_entrada_vigencia;
            referencia = config.referencia;
            condiciones_de_venta = config.condiciones_de_venta;
            condiciones_de_pago = config.condiciones_de_pago;
            fecha_entrega = config.fecha_entrega;
            unidad_medida = config.unidad_medida;
            mail_envio_cotizacion_subject = config.mail_envio_cotizacion_subject;
            mail_envio_cotizacion_text = config.mail_envio_cotizacion_text;
            public_logo = config.public_logo;
            agendamiento_mail = config.agendamiento_mail;
            cotizacion_mail = config.cotizacion_mail;
            mail_seguimiento_cotizacion = config.mail_seguimiento_cotizacion;
            mail_post_venta = config.mail_post_venta;
            mail_alerta_fase_2 = config.mail_alerta_fase_2;
            datos_transferencia_nombre = config.datos_transferencia_nombre;
            datos_transferencia_banco = config.datos_transferencia_banco;
            datos_transferencia_cta = config.datos_transferencia_cta;
            datos_transferencia_rut = config.datos_transferencia_rut;
            datos_transferencia_email = config.datos_transferencia_email;
            webPage = config.webPage;
            email_contacto = config.email_contacto;
            nota_descripcion = config.nota_descripcion;
            observacion = config.observacion;

            return true;
        }
    }
}