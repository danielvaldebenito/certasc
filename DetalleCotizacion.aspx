<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DetalleCotizacion.aspx.cs" Inherits="DetalleCotizacion" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Certasc</title>
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
    <link href="bower_components/jquery.steps/demo/css/jquery.steps.css" rel="stylesheet" />
    <link href="bower_components/bootstrap-switch.css" rel="stylesheet" />
    <link href="bower_components/magnific-popup.css" rel="stylesheet" />
    <link href="bower_components/jquery-te-1.4.0.css" rel="stylesheet" />
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/detalleCotizacion.css" rel="stylesheet" />
    <link href="Content/bootstrap-datetimepicker.css" rel="stylesheet" />
    <style>
        .list-group-item {
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
        }
    </style>
</head>
<body>

    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="page-header">
                    <h1 id="titleNumeroCotizacion"></h1>
                </div>
            </div>
        </div>

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
                                <label>Estado</label>
                                <span id="rg_estado"></span>
                            </li>
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
                            <li class="list-group-item list-group-item-flex">
                                <label>Valor</label>
                                <span id="r_valor"></span>
                            </li>
                            <li class="list-group-item list-group-item-danger list-group-item-flex">
                                <label>Descuento</label>
                                <span id="r_descuento"></span>
                            </li>
                            <li class="list-group-item list-group-item-warning list-group-item-flex">
                                <label>Recargo</label>
                                <span id="r_recargo"></span>
                            </li>
                            <li class="list-group-item list-group-item-success list-group-item-flex">
                                <label>Total</label>
                                <span id="r_total" style="text-align: right"></span>
                            </li>
                        </ul>
                    </div>
                    <div class="col-xs-12 col-sm-6 col-md-3 col-lg-3"></div>

                    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                        <div class="panel panel-default">
                            <div class="panel-heading">
                                <h3>HITOS
                                    <span id="send-mail" title="Enviar Correo"><i class="fa fa-envelope-o"></i></span>
                                    <span id="add-hito" title="Agregar Hito"><i class="fa fa-plus"></i></span>
                                    <span id="reenviar" title="Reenviar Cotización">
                                        <i class="fa fa-refresh"></i>
                                    </span>
                                </h3>
                            </div>
                            <div class="panel-body" data-height-full="true">
                                <div class="row">
                                    <table id="grid"></table>
                                    <div id="pager"></div>
                                </div>

                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </div>
    </div>

    <div id="add-hito-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-hito-form">

                        <div class="form-group col-xs-12">
                            <label for="add-tipo-hito">Tipo de Hito (*)</label>
                            <select class="form-control" id="add-tipo-hito" required>
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="add-obs">Observaciones (*)</label>
                            <textarea id="add-obs" placeholder="Observaciones" class="form-control" required></textarea>
                        </div>

                        <div class="form-group col-xs-12" id="divIngeniero">
                            <label for="add-ingeniero">Ingeniero (*)</label>

                            <select class="form-control" id="add-ingeniero">
                                <option value="0">Seleccione...</option>
                            </select>


                        </div>

                        <div class="form-group col-xs-12" id="divFechaCompromiso">
                            <label for="add-fechaCompromiso">Fecha Compromiso (*)</label>
                            <div class="input-group">
                                <input type="text" class="form-control" id="add-fechaCompromiso" placeholder="Fecha Compromiso" />
                                <span id="ShowCalendar" class="input-group-addon btn btn-default">
                                    <i class="fa fa-calendar"></i>
                                </span>
                            </div>
                        </div>


                    </form>
                </div>
            </div>
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
                            <textarea class="form-control" id="correo" rows="8"></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <div id="reenviar-dialog" class="dialog">
        <form id="reenviar-form">
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
                <textarea class="form-control" id="correo1" rows="5"></textarea>
            </div>
        </form>

    </div>

    <div class="lock-screen dialog">
        <i class="fa fa-cog fa-spin fa-3x fa-fw margin-bottom"></i>
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
<script src="Scripts/moment-with-locales.js"></script>
<script src="Scripts/bootstrap-datetimepicker.js"></script>
<script src="bower_components/jquery-te-1.4.0.min.js"></script>
<script src="js/menu.js"></script>
<script src="js/certel.js"></script>
<script src="js/jquery.Rut.js"></script>
<script src="js/elementos/detalleCotizacion.js"></script>
</html>
