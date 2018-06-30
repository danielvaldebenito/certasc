<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Servicios.aspx.cs" Inherits="Servicios" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Certasc</title>
    <link rel="shortcut icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link rel="icon" href="css/images/favicon.ico" type="image/x-icon" />
    <link href="bower_components/font-awesome/css/font-awesome.min.css" rel="stylesheet" />
    <link href="bower_components/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/jquery-ui/jquery-ui.min.css" rel="stylesheet"/>
    <link href="bower_components/jquery-ui/jquery-ui.theme.min.css" rel="stylesheet" />
    <link href="bower_components/jqGrid/ui.jqgrid.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/alertify.min.css" rel="stylesheet" />
    <link href="bower_components/alertify.js/bootstrap.min.css" rel="stylesheet" />
    <link href="bower_components/select2/dist/css/select2.min.css" rel="stylesheet" />
    <link href="bower_components/jquery.steps.css" rel="stylesheet" />
    <link href="css/select2.min.css" rel="stylesheet"/>
    <link href="css/certel.css" rel="stylesheet" />
    
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="page-header">
                    <h1>Servicios</h1>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-3 col-lg-3">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>FILTROS <span id="f_remove"><i class="fa fa-remove"></i></span></h3>
                    </div>
                    <div class="panel-body">
                        <form id="formFiltros">
                            <div class="form-group">
                                <label for="f_cliente">Cliente</label>
                                <input type="text" class="form-control" id="f_cliente"/>
                            </div>
                            <div class="form-group">
                                <label for="f_cliente">Fecha</label>
                                <div class="input-group">
                                    <input type="text" class="form-control" id="f_desde" placeholder="Desde" readonly/>
                                    <span class="input-group-addon"><i class="fa fa-calendar"></i></span>
                                    <input type="text" class="form-control" id="f_hasta" placeholder="Hasta" readonly/>
                                </div>
                                
                            </div>

                        </form>
                    </div>
                </div>
            </div>
            <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>REGISTROS <span id="add-service"><i class="fa fa-plus"></i></span></h3>
                    </div>
                    <div class="panel-body" data-height-full="true">
                        <table id="grid"></table>
                        <div id="pager"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <div class="foot-page">Certel - Certificación en Elevación S.A.</div>

    <div id="add-service-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    
                    <form id="add-service-form">
                        <div class="form-group col-xs-12">
                            <label for="add-cliente">Cliente (*)</label>
                            <div class="input-group">
                                <select class="form-control" id="add-cliente" required>
                                    <option value="0">Seleccione...</option>
                                </select>
                                <span id="add-client" class="input-group-addon btn btn-default">
                                    <i class="fa fa-plus-circle"></i>
                                </span>
                                 
                            </div>
                                
                        </div>
                       
                        
                        <div class="form-group col-xs-12">
                            <label for="add-it">IT (*)</label>
                            <input id="add-it" type="text" class="form-control" placeholder="IT" required />
                        </div>
                    </form>
                </div>
            </div>
        </div>       
    </div>
    <div id="add-client-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-md-12 col-sm-12 col lg-12">
                    <form id="add-client-form">
                        <div class="form-group">
                            <label for="ac_rut">Rut (*)</label>
                            <input type="text" value="" class="form-control" id="ac_rut" required maxlength="50" />
                        </div>
                        <div class="form-group">
                            <label for="ac_nombre">Nombre (*)</label>
                            <input type="text" value="" class="form-control" id="ac_nombre" required maxlength="50"/>
                        </div>
                        <div class="form-group">
                            <label for="ac_direccion">Dirección</label>
                            <input type="text" value="" class="form-control" id="ac_direccion" maxlength="200" />
                        </div>
                        <div class="form-group">
                            <label for="ac_telefono">Teléfono Contacto</label>
                            <input type="text" value="" class="form-control" id="ac_telefono" />
                        </div>
                        <div class="form-group">
                            <label for="ac_mail">Email Contacto</label>
                            <input type="email" value="" class="form-control" id="ac_mail" />
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
    <div id="inspecciones-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                    <button class="btn btn-default" id="add-inspeccion">
                        <i class="fa fa-plus"></i>
                        NUEVA INSPECCIÓN
                    </button>
                    
                    <table id="grid-inspecciones"></table>
                    <div id="pager-inspecciones"></div>
                </div>
            </div>
        </div>
    </div>

    <%--Add Inspecciones--%>
    <%--Add Inspecciones--%>
    <div id="add-inspeccion-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                    <form id="add-inspeccion-form">
                        <div class="col-xs-12 col-lg-12 col-md-12 col-sm-12">
                            <div class="form-group">
                                <label for="ai-it">IT (*)</label>
                                <input id="ai-it" type="text" class="form-control" placeholder="IT Servicio" required maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="ai-ubicacion">Ubicación (*)</label>
                                <input id="ai-ubicacion" type="text" class="form-control" placeholder="Ubicación" required maxlength="300" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-6 col-md-6 col-sm-12">
                            <div class="form-group">
                                <label for="ai-edificio">Nombre del edificio (*)</label>
                                <input id="ai-edificio" type="text" class="form-control" placeholder="Nombre del edificio" required maxlength="300" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-fecha-instalacion">Fecha Instalación (*)</label>
                                <input id="ai-fecha-instalacion" type="text" class="form-control" readonly required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-fecha-inspeccion">Fecha Inspección (*)</label>
                                <input id="ai-fecha-inspeccion" type="text" class="form-control" readonly required />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-aparato">Tipo de elevador (*)</label>
                                <select class="form-control" id="ai-aparato" required></select>
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-tipo-funcionamiento">Tipo de funcionamiento (*)</label>
                                <select class="form-control" id="ai-tipo-funcionamiento" required></select>
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-4 col-md-4 col-sm-12">
                            <div class="form-group">
                                <label for="ai-numero">Número único del elevador (*)</label>
                                <input id="ai-numero" type="text" class="form-control" required" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-8 col-md-8 col-sm-12">
                            <div class="form-group">
                                <label for="ai-nombre">Nombre del proyecto (*)</label>
                                <input id="ai-nombre" type="text" class="form-control" required maxlength="50" placeholder="Nombre del proyecto" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-4 col-md-4 col-sm-12">
                            <div class="form-group">
                                <label for="ai-destino">Destino proyecto (*)</label>
                                <select id="ai-destino" class="form-control" required>
                                    <option value="0">Seleccione...</option>
                                </select>
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-permiso-edificacion">Permiso edificación</label>
                                <input id="ai-permiso-edificacion" type="text" class="form-control" maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-3 col-md-3 col-sm-12">
                            <div class="form-group">
                                <label for="ai-recepcion-municipal">Recepción Municipal</label>
                                <input id="ai-recepcion-municipal" type="text" class="form-control" maxlength="50" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-2 col-md-2 col-sm-12">
                            <div class="form-group">
                                <label for="ai-altura">Altura en pisos</label>
                                <input id="ai-altura" type="number" class="form-control" min="1" />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-2 col-md-2 col-sm-12">
                            <div class="form-group">
                                <label for="ai-ingeniero">Ingeniero Inspector(*) <br /></label>
                                <select class="form-control" id="ai-ingeniero" required></select>
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-5 col-md-5 col-sm-12">
                            <div class="form-group">
                                <label for="ai-fec">Fecha de Emisión del Certificado de Inspección de Certificación</label>
                                <input type="text" class="form-control" id="ai-fec" readonly />
                            </div>
                        </div>
                        <div class="col-xs-12 col-lg-5 col-md-5 col-sm-12">
                            <div class="form-group">
                                <label for="ai-fvc">Fecha de Vencimiento del Certificado de Inspección de Certificación</label>
                                <input type="text" class="form-control" id="ai-fvc" readonly />
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>    
    
   

</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery.steps.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="js/menu.js?20062017"></script>
<script src="js/select2.full.min.js"></script>
<script src="js/jquery.Rut.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/servicios.js?20062017"></script>
</html>
