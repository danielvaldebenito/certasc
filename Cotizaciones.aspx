<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Cotizaciones.aspx.cs" Inherits="Cotizaciones" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Certasc - Cotizaciones</title>
    <link rel="shortcut icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link rel="icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link href="bower_components/font-awesome/css/font-awesome.min.css" rel="stylesheet" />
    <link href="bower_components/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.theme.min.css" rel="stylesheet" />
    <link href="bower_components/jqGrid/ui.jqgrid.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/alertify.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/select2/dist/css/select2.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-te-1.4.0.css" rel="stylesheet" />
    <link href="bower_components/bootstrap-switch.css" rel="stylesheet" />
    <link href="bower_components/magnific-popup.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
    <link href="Content/bootstrap-datetimepicker.css" rel="stylesheet" />
    <link href="css/elementos/cotizaciones.css?04042017" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>COTIZACIONES 
                            <span id="f_remove"><i class="fa fa-remove"></i></span>
                            <span id="add"><i class="fa fa-plus"></i>NUEVA</span>
                        </h3>
                    </div>
                    <div class="panel-body" data-height-full="true">
                        <div class="row">
                            <form id="formFiltros">
                                <div class="form-group col-xs-12 col-sm-12 col-md-3 col-lg-2">
                                    <input type="text" class="form-control" id="f_name" placeholder="Nombre" />
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-3 col-lg-3">
                                    <div class="input-group">
                                        <input type="text" id="f_inicio" placeholder="Desde" class="form-control" />
                                        <span class="input-group-addon">
                                            <i class="fa fa-calendar"></i>
                                        </span>
                                        <input type="text" id="f_fin" placeholder="Hasta" class="form-control" />
                                    </div>
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-2 col-lg-2">
                                    <select class="form-control" id="f_estado"></select>
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-2 col-lg-2">
                                    <input type="checkbox" value="" id="f_alertF2" />
                                    <label title="Mostrar sólo cotizacines para alarmar próxima Fase II"><i class="fa fa-bell fa-2x" style="color: #F44336"></i></label>
                                </div>
                                <div class="form-group col-xs-12 col-sm-12 col-md-3 col-lg-3">

                                    <button class="btn btn-default btn-block" type="submit">
                                        <i class="fa fa-search"></i>
                                        Buscar
                                    </button>
                                </div>
                            </form>
                        </div>

                        <div class="row">
                            <table id="grid"></table>
                            <div id="pager"></div>
                        </div>

                    </div>
                    <div class="panel-footer">
                        <div class="row">
                            <div class="col-xs-12 col-sm-4 col-md-4 col-lg-4 text-center">
                                UF: <span id="uf"></span>
                            </div>
                            <div class="col-xs-12 col-sm-4 col-md-4 col-lg-4 text-center">
                                Dolar: <span id="dolar"></span>
                            </div>
                            <div class="col-xs-12 col-sm-4 col-md-4 col-lg-4 text-center">
                                Euro: <span id="euro"></span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div id="add-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">

                    <div id="rootwizard" class="tabbable tabs-left">
                        <ul>
                            <li><a href="#tab1" data-toggle="tab">Cliente</a></li>
                            <li><a href="#tab2" data-toggle="tab">Datos Generales</a></li>
                            <li><a href="#tab3" data-toggle="tab">Servicios</a></li>
                            <li><a href="#tab7" data-toggle="tab">Costos Logísticos</a></li>
                            <%--<li><a href="#tab6" data-toggle="tab">Envío</a></li>--%>
                            <li><a href="#tab5" data-toggle="tab">Resumen</a></li>
                        </ul>
                        <div class="tab-content">
                            <div class="tab-pane" id="tab1">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>Cliente</h3>
                                    </div>
                                    <div class="panel-body">
                                        <form id="form-cliente">
                                            <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-3">
                                                <label>Rut:</label>
                                                <div class="input-group">
                                                    <input type="text" name="rut-cliente" id="rut-cliente" value="" required class="form-control" />
                                                    <span class="input-group-addon btn btn-default" id="search-client">
                                                        <i class="fa fa-search"></i>
                                                    </span>
                                                </div>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-8 col-lg-9">
                                                <label>Nombre:</label>
                                                <input type="text" name="nombre-cliente" id="nombre-cliente" required value="" class="form-control" />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-6 col-lg-6">
                                                <label>Giro:</label>
                                                <input type="text" name="giro-cliente" id="giro-cliente" value="" class="form-control" />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-6 col-lg-6">
                                                <label>Dirección:</label>
                                                <input type="text" name="direccion-cliente" id="direccion-cliente" value="" class="form-control" />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-4 col-md-4 col-lg-4">
                                                <label>Nombre Contacto:</label>
                                                <input type="text" name="contacto-cliente" id="contacto-cliente" required value="" class="form-control" />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-4 col-md-4 col-lg-4">
                                                <label>E-mail contacto:</label>
                                                <input type="email" name="email-contacto-cliente" id="email-contacto-cliente" required value="" class="form-control" />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-4 col-md-4 col-lg-4">
                                                <label>Teléfono contacto:</label>
                                                <input type="text" name="fono-contacto-cliente" id="fono-contacto-cliente" required value="" class="form-control" />
                                            </div>


                                        </form>
                                    </div>
                                </div>

                            </div>
                            <div id="tab2" class="tab-pane">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>Datos generales</h3>
                                    </div>
                                    <div class="panel-body">
                                        <form id="form-generales">
                                            <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                                                <label>Fecha Doc:</label>
                                                <input type="text" name="fecha_doc" id="fecha_doc" readonly value="" required class="form-control" />
                                            </div>

                                            <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                                                <label>Fuente Solicitud:</label>
                                                <select name="fuente_solicitud" id="fuente_solicitud" required class="form-control">

                                                    <option value="EMAIL">EMAIL</option>
                                                    <option value="TELEFÓNICA">TELEFÓNICA</option>
                                                    <option value="OTRO">OTRO</option>
                                                </select>
                                            </div>

                                            <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                                                <label>Forma de pago:</label>
                                                <select name="forma_pago" id="forma_pago" required class="form-control">
                                                </select>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                                                <label>Moneda:</label>
                                                <select id="moneda" name="moneda" required class="form-control">
                                                </select>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                                                <label>Vendedor:</label>
                                                <select name="vendedor" id="vendedor" required class="form-control">
                                                    <option value="0">SELECCIONE...</option>
                                                </select>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                                                <label>Ciudad:</label>
                                                <select id="ciudad" required class="form-control">
                                                    <option value="0">SELECCIONE...</option>
                                                </select>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                                                <label>Fecha validez:</label>
                                                <input type="text" class="form-control" id="validez" required readonly />
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <label>Observación:</label>
                                                <textarea rows="2" name="observacion" id="observacion" class="form-control" maxlength="500"></textarea>
                                            </div>
                                            <div class="form-group col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <label>Nota:</label>
                                                <textarea rows="2" name="nota" id="nota" class="form-control" maxlength="500"></textarea>
                                            </div>
                                        </form>
                                    </div>

                                </div>

                            </div>
                            <div class="tab-pane" id="tab3">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>Servicios <span id="add-item-button"><i class="fa fa-plus-circle"></i>Agregar ítem </span></h3>
                                    </div>
                                    <div class="panel-body">

                                        <table id="grid-items" class="table table-responsive table-bordered table-condensed table-hover">
                                            <thead>
                                                <tr>
                                                    <th rowspan="2">Cantidad</th>
                                                    <th rowspan="2">Descripción</th>
                                                    <th rowspan="2">Ubicación</th>
                                                    <th rowspan="2">Valor U.</th>
                                                    <th colspan="2">Descuento</th>
                                                    <th rowspan="2">Precio Cliente</th>
                                                    <th rowspan="2"></th>
                                                </tr>
                                                <tr>
                                                    <th>Tipo</th>
                                                    <th>Valor</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                            </tbody>
                                            <tfoot>
                                                <tr>
                                                    <th class="text-right" id="grid-items-cantidad"></th>
                                                    <th></th>
                                                    <th></th>
                                                    <th class="text-right" id="grid-items-valor"></th>
                                                    <th></th>
                                                    <th class="text-right" id="grid-items-descuento"></th>
                                                    <th class="text-right" id="grid-items-total"></th>
                                                </tr>
                                            </tfoot>
                                        </table>
                                    </div>

                                </div>

                                <div>
                                </div>
                            </div>
                            <div id="tab7" class="tab-pane">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>Costos Logísticos</h3>
                                    </div>
                                    <div class="panel-body">
                                        <form id="gastos-form">
                                            <div class="input-group">
                                                <span class="input-group-addon">#</span>
                                                <input id="gastos-cantidad" type="number" class="form-control" value="1" min="1" required />
                                                <span class="input-group-addon">Tipo</span>
                                                <select class="form-control" id="tipo-gasto" required></select>
                                                <span class="input-group-addon">$</span>
                                                <input type="number" value="" placeholder="Valor" class="form-control" readonly id="valor-gasto" required />
                                                <span id="edit-valor-gasto" class="input-group-addon btn btn-default">
                                                    <i class="fa fa-pencil"></i>
                                                </span>
                                                <span id="add-gasto" class="input-group-addon btn btn-default">
                                                    <i class="fa fa-plus-circle"></i>Agregar
                                                </span>
                                            </div>
                                        </form>
                                        <table id="gastos-table" class="table table-responsive table-bordered table-condensed table-hover">

                                            <thead>
                                                <tr>
                                                    <th>#</th>
                                                    <th>Ítem</th>
                                                    <th>Valor</th>
                                                    <th>Total</th>
                                                    <th></th>
                                                </tr>
                                            </thead>
                                            <tbody></tbody>
                                            <tfoot>
                                                <tr>
                                                    <th></th>
                                                    <th>TOTAL</th>
                                                    <th></th>
                                                    <th class="text-right" id="total-gastos"></th>
                                                    <th></th>
                                                </tr>
                                            </tfoot>
                                        </table>
                                    </div>
                                </div>
                            </div>

                            <div class="tab-pane" id="tab5">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>Resumen</h3>
                                    </div>
                                    <div class="panel-body">
                                        <div class="row">
                                            <div class="col-xs-12 col-sm-6 col-md-6 col-lg-6">
                                                <h3>CLIENTE</h3>
                                                <ul class="list-group">
                                                    <li class="list-group-item">
                                                        <label>Rut</label>
                                                        <span id="rc_rut"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Nombre</label>
                                                        <span id="rc_nombre"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Giro</label>
                                                        <span id="rc_giro"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Dirección</label>
                                                        <span id="rc_direccion"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Contacto</label>
                                                        <span id="rc_contacto"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Email</label>
                                                        <span id="rc_email"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Fono</label>
                                                        <span id="rc_fono"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Ciudad</label>
                                                        <span id="rc_ciudad"></span>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="col-xs-12 col-sm-6 col-md-6 col-lg-6">
                                                <h3>DATOS GENERALES</h3>
                                                <ul class="list-group">
                                                    <li class="list-group-item">
                                                        <label>Fecha Doc</label>
                                                        <span id="rg_fecha_doc"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Fuente Solicitud</label>
                                                        <span id="rg_fuente_solicitud"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Forma de Pago</label>
                                                        <span id="rg_forma_de_pago"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Moneda</label>
                                                        <span id="rg_moneda"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Vendedor</label>
                                                        <span id="rg_vendedor"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Fecha Validez</label>
                                                        <span id="rg_validez"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Observaciones</label>
                                                        <span id="rg_observaciones"></span>
                                                    </li>
                                                    <li class="list-group-item">
                                                        <label>Nota</label>
                                                        <span id="rg_nota"></span>
                                                    </li>
                                                </ul>

                                            </div>
                                            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                                                <table id="grid-items-resume" class="table table-bordered table-condensed table-hover">
                                                    <caption>
                                                        <h3>SERVICIOS</h3>
                                                    </caption>
                                                    <thead>
                                                        <tr>
                                                            <th rowspan="2">Cantidad</th>
                                                            <th rowspan="2">Descripción</th>
                                                            <th rowspan="2">Ubicación</th>
                                                            <th rowspan="2">Valor U.</th>
                                                            <th colspan="2">Descuento</th>
                                                            <th rowspan="2">Precio Cliente</th>
                                                        </tr>
                                                        <tr>
                                                            <th>Tipo</th>
                                                            <th>Valor</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                    </tbody>
                                                    <tfoot>
                                                        <tr>
                                                            <th class="text-right" id="grid-items-resume-cantidad"></th>
                                                            <th></th>
                                                            <th></th>
                                                            <th class="text-right" id="grid-items-resume-valor"></th>
                                                            <th></th>
                                                            <th class="text-right" id="grid-items-resume-descuento"></th>
                                                            <th class="text-right" id="grid-items-resume-total"></th>
                                                        </tr>
                                                    </tfoot>
                                                </table>
                                            </div>
                                            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                                                <table id="gastos-table-resume" class="table table-responsive table-bordered table-condensed table-hover">
                                                    <caption>
                                                        <h3>COSTOS LOGÍSTICOS</h3>
                                                    </caption>
                                                    <thead>
                                                        <tr>
                                                            <th>#</th>
                                                            <th>Ítem</th>
                                                            <th>Valor</th>
                                                            <th>Total</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody></tbody>
                                                    <tfoot>
                                                        <tr>
                                                            <th></th>
                                                            <th>TOTAL</th>
                                                            <th></th>
                                                            <th class="text-right" id="total-gastos-resume"></th>

                                                        </tr>
                                                    </tfoot>
                                                </table>
                                            </div>
                                            <div class="col-xs-12">
                                                <h3>VALORES</h3>
                                                <ul class="list-group">
                                                    <li class="list-group-item">
                                                        <label>Valor</label>
                                                        <span id="r_valor"></span>
                                                    </li>
                                                    <li class="list-group-item list-group-item-danger">
                                                        <label>Descuento</label>
                                                        <span id="r_descuento"></span>
                                                    </li>
                                                    <li class="list-group-item list-group-item-warning">
                                                        <label>Recargo</label>
                                                        <span id="r_recargo"></span>
                                                    </li>
                                                    <li class="list-group-item list-group-item-success">
                                                        <label>Total</label>
                                                        <span id="r_total"></span>
                                                    </li>
                                                </ul>
                                            </div>
                                            <div class="col-xs-12 col-sm-6 col-md-3 col-lg-3"></div>
                                        </div>
                                    </div>

                                </div>

                            </div>
                            <%--<div id="tab6" class="tab-pane">
                                <div class="panel panel-default">
                                    <div class="panel-heading">
                                        <h3>ENVÍO</h3>
                                    </div>
                                    <div class="panel-body">
                                        <div class="container">
                                            <div class="row">
                                                <div class="col-xs-12">
                                                    <input type="checkbox" value="" id="no_send_mail" />
                                                    No enviar correo
                                                    <form id="mail">
                                                        <div class="form-group">
                                                            <label>Asunto</label>
                                                            <input type="text" id="asunto" class="form-control" placeholder="Asunto" />
                                                        </div>
                                                        <div class="form-group" id="para-fg">
                                                            <label>Para</label>
                                                            <div class="input-group">
                                                                <input type="email" id="para" name="para" class="form-control para" placeholder="Para" />
                                                                <span class="input-group-addon btn btn-default" id="add-para">
                                                                    <i class="fa fa-plus"></i>
                                                                </span>
                                                            </div>
                                                        </div>
                                                        <div class="form-group" id="cc-fg">
                                                            <label>CC</label>
                                                            <div class="input-group">
                                                                <input type="email" id="cc" name="cc" class="form-control cc" placeholder="Con copia" />
                                                                <span class="input-group-addon btn btn-default" id="add-cc">
                                                                    <i class="fa fa-plus"></i>
                                                                </span>
                                                            </div>
                                                        </div>
                                                        <div class="form-group" id="cco-fg">
                                                            <label>CCO</label>
                                                            <div class="input-group">
                                                                <input type="email" id="cco" name="cco" class="form-control cco" placeholder="Con copia oculta" />
                                                                <span class="input-group-addon btn btn-default" id="add-cco">
                                                                    <i class="fa fa-plus"></i>
                                                                </span>
                                                            </div>
                                                        </div>
                                                        <div class="form-group">
                                                            <textarea class="form-control" id="correo" rows="8"></textarea>
                                                        </div>
                                                    </form>
                                                    <div class="panel panel-default">
                                                        <div class="panel-heading">
                                                            <h3>ARCHIVOS ADJUNTOS
                                                                <span id="add-file"><i class="fa fa-upload"></i></span>
                                                            </h3>
                                                        </div>
                                                        <ul id="files" class="list-group"></ul>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>--%>
                            <br />
                            <ul class="pager wizard">
                                <li class="previous first"><a href="javascript:;"><i class="fa fa-fast-backward"></i></a></li>
                                <li class="previous"><a href="javascript:;"><i class="fa fa-backward"></i>ANTERIOR</a></li>
                                <li class="next last"><a href="javascript:;"><i class="fa fa-fast-forward"></i></a></li>
                                <li class="next"><a href="javascript:;"><i class="fa fa-forward"></i>SIGUIENTE</a></li>
                                <li class="finish pull-right"><a href="javascript:;"><i class="fa fa-check"></i>FINALIZAR</a></li>
                            </ul>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    </div>
    <div id="clientes-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <div class="form-group col-xs-6">
                        <label>Rut</label>
                        <input type="text" value="" id="fc_rut" class="form-control" />
                    </div>
                    <div class="form-group col-xs-6">
                        <label>Nombre</label>
                        <input type="text" value="" id="fc_nombre" class="form-control" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12">
                    <table id="grid-clientes"></table>
                    <div id="pager-clientes"></div>
                </div>
            </div>
        </div>
    </div>
    <div id="productos-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="form-group col-xs-12">
                    <label>Descripción</label>
                    <input type="text" value="" id="fp_nombre" class="form-control" />
                </div>
            </div>
            <div class="row">
                <div class="col-xs-12">
                    <table id="grid-productos"></table>
                    <div id="pager-productos"></div>
                </div>
            </div>
        </div>
    </div>

    <div class="dialog" id="add-item-dialog">
        <form id="form-detalle">
            <div class="container-fluid">
                <div class="row">
                    <div class="form-group col-xs-12 col-sm-4 col-md-3 col-lg-3">
                        <label>Producto (*)</label>
                        <select id="producto" class="form-control" required></select>
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-7 col-lg-7">
                        <label>Ubicación</label>
                        <input type="text" id="ubicacion" class="form-control" maxlength="100" />
                    </div>

                    <div class="form-group col-xs-12 col-sm-2 col-md-2 col-lg-2">
                        <label>Cantidad (*)</label>
                        <input type="number" id="item_cantidad" value="1" class="form-control" required />
                    </div>

                    <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                        <label>Elevador</label>
                        <select class="form-control" id="tipo_elevador" required></select>
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                        <label>Funcionamiento</label>
                        <select class="form-control" id="tipo_funcionamiento"></select>
                    </div>

                    <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                        <div class="form-group">
                            <label for="altura">N° Paradas</label>
                            <input id="altura" type="number" class="form-control" value="2" min="0" required />
                        </div>
                    </div>


                    <div class="form-group col-xs-12 col-sm-6 col-md-3 col-lg-3">
                        <label>Año Instalación</label>
                        <select id="ano_instalacion" class="form-control"></select>
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                        <label>Uso</label>
                        <select id="uso" class="form-control">
                        </select>
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                        <label>Marca</label>
                        <select class="form-control" id="marca"></select>
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                        <label>Empresa Instaladora</label>
                        <input type="text" id="empresa_instaladora" class="form-control" maxlength="100" />
                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-8 col-lg-12">
                        <label>Descripción (*)</label>
                        <div class="input-group">
                            <textarea id="item_descripcion" rows="2" readonly class="form-control" required maxlength="500"></textarea>
                            <span class="input-group-addon" id="edit-item-description" title="Editar Descripción">
                                <i class="fa fa-pencil"></i>
                            </span>
                        </div>

                    </div>

                    <div class="form-group col-xs-12 col-sm-4 col-md-4 col-lg-4">
                        <label>Unitario (*)</label>
                        <div class="input-group">
                            <input type="number" id="item_unitario" value="0" class="form-control" required readonly />
                            <span id="edit-price" class="input-group-addon">
                                <i class="fa fa-pencil"></i>
                            </span>
                        </div>

                    </div>
                    <div class="form-group col-xs-12 col-sm-6 col-md-4 col-lg-4">
                        <label>Descuento (*)</label>
                        <div class="input-group">
                            <select id="item_tipo_descuento" class="form-control input-group-addon">
                                <option value="0">-</option>
                                <option value="VALOR">$€UF</option>
                                <option value="PORCENTAJE">%</option>
                            </select>
                            <span class="input-group-addon"><i class="fa fa-usd"></i></span>
                            <input type="number" id="item_descuento" disabled value="0" class="form-control input-group-addon" required />
                        </div>
                    </div>
                    <div class="form-group col-xs-12 col-sm-4 col-md-4 col-lg-4">
                        <label>Total</label>
                        <input type="number" id="item_total" value="0" class="form-control" readonly />
                    </div>
                </div>
            </div>

        </form>
    </div>
    <div id="add-file-dialog" class="dialog">
        <form id="add-file-form">
            <div class="form-group">
                <label>Nombre del archivo</label>
                <input type="text" id="add-file-name" class="form-control" value="" required />
            </div>
            <div class="form-group">
                <label>Nombre del archivo</label>
                <input type="file" id="add-file-file" class="form-control" value="" required />
            </div>
        </form>
    </div>
    <div id="add-select-dialog" class="dialog">
        <div class="form-group">
            <button class="btn btn-default btn-block" id="desde-cero">Desde Cero</button>
        </div>
        <div class="form-group">
            <div class="input-group">
                <input type="text" class="form-control" value="" id="a-partir-de-it" placeholder="A partir del IT..." />
                <span class="input-group-addon btn btn-default" id="a-partir-de">Comenzar</span>
            </div>
        </div>

    </div>
    <div id="services-dialog" class="dialog">
        <div id="chart" style="height: 200px"></div>
        <table id="services-grid"></table>
        <div id="services-pager"></div>
        <div class="row">
            <div class="col-xs-12">
                <button class="btn btn-primary btn-block" id="agendar-fase2">Agendar Fase II</button>
            </div>
            <div class="col-xs-12">
                <button class="btn btn-primary btn-block" id="post-venta" style="display: none">Enviar E-mail Post-Venta</button>
            </div>
        </div>
    </div>
    <div id="fase2-dialog" class="dialog">
        <div class="form-group">
            <label>Fecha para Fase 2</label>
            <input type="text" value="" id="fecha-fase2" class="form-control" maxlength="16" placeholder="Ej. 24-03-2018 12:00" />
        </div>
    </div>
    <div id="send-mail-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="send-mail-form">
                        <div class="form-group">
                            <input type="text" id="mail-asunto" class="form-control" placeholder="Asunto" required />
                        </div>
                        <div class="form-group">
                            <input type="email" id="mail-destinatario" class="form-control" placeholder="Para" required />
                        </div>
                        <div class="form-group">
                            <textarea class="form-control" id="correo2" rows="8"></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
    <div id="send-alert-fase2-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="send-alert-fase2-form">
                        <div class="form-group">
                            <label>Asunto</label>
                            <input type="text" id="f2-asunto" class="form-control" placeholder="Asunto" required />
                        </div>
                        <div class="form-group">
                            <label>Para</label>
                            <input type="email" id="f2-destinatario" class="form-control" placeholder="Para" required />
                        </div>
                        <div class="form-group">
                            <textarea class="form-control" id="f2-correo" rows="4"></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
    <div id="send-mail-cot-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h3>ENVÍO</h3>
                        </div>
                        <div class="panel-body">

                            <form id="send-mail-cot-form">
                                <div class="form-group">
                                    <label>Asunto</label>
                                    <input type="text" id="asunto" class="form-control" placeholder="Asunto" />
                                </div>
                                <div class="form-group" id="para-fg">
                                    <label>Para</label>
                                    <div class="input-group">
                                        <input type="email" id="para" name="para" class="form-control para" placeholder="Para" />
                                        <span class="input-group-addon btn btn-default" id="add-para">
                                            <i class="fa fa-plus"></i>
                                        </span>
                                    </div>
                                </div>
                                <div class="form-group" id="cc-fg">
                                    <label>CC</label>
                                    <div class="input-group">
                                        <input type="email" id="cc" name="cc" class="form-control cc" placeholder="Con copia" />
                                        <span class="input-group-addon btn btn-default" id="add-cc">
                                            <i class="fa fa-plus"></i>
                                        </span>
                                    </div>
                                </div>
                                <div class="form-group" id="cco-fg">
                                    <label>CCO</label>
                                    <div class="input-group">
                                        <input type="email" id="cco" name="cco" class="form-control cco" placeholder="Con copia oculta" />
                                        <span class="input-group-addon btn btn-default" id="add-cco">
                                            <i class="fa fa-plus"></i>
                                        </span>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <textarea class="form-control" id="correo" rows="8"></textarea>
                                </div>
                            </form>
                            <div class="panel panel-default">
                                <div class="panel-heading">
                                    <h3>ARCHIVOS ADJUNTOS
                                         <span id="add-file"><i class="fa fa-upload"></i></span>
                                    </h3>
                                </div>
                                <ul id="files" class="list-group"></ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>


    <div class="foot-page">Certel - Certificación en Elevación S.A.</div>
    <%--Lock Screen--%>
    <div class="lock-screen">
        <i class="fa fa-cog fa-spin fa-5x fa-fw margin-bottom"></i>
        <span class="sr-only">Cargando...</span>
    </div>
</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="bower_components/bootstrap-switch.js"></script>
<script src="bower_components/magnific-popup.js"></script>
<script src="bower_components/twitter-bootstrap-wizard/jquery.bootstrap.wizard.js"></script>
<script src="bower_components/jquery-te-1.4.0.min.js"></script>
<script src="Scripts/moment-with-locales.js"></script>
<script src="Scripts/bootstrap-datetimepicker.js"></script>
<script src="bower_components/linq.min.js"></script>
<script src="js/menu.js?20062017"></script>
<script src="js/jquery.Rut.js"></script>

<script src="https://code.highcharts.com/highcharts.js"></script>
<script src="js/certel.js?29012018"></script>

<script src="js/elementos/cotizaciones.js?140420182013"></script>
</html>
