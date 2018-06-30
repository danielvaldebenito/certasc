<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CostosLogisticos.aspx.cs" Inherits="CostosLogisticos" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Certasc - Costos Logísticos</title>
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
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/costoslogisticos.css?04042017" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-3 col-lg-3">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>FILTROS
                            <span id="f_remove"><i class="fa fa-remove"></i></span>
                        </h3>
                    </div>
                    <div class="panel-body">
                        <form id="filter-form">
                            <div class="form-group">
                                <label>Ciudad</label>
                                <select id="f_ciudad" class="form-control"></select>
                            </div>
                            <div class="form-group">
                                <label>Tipo Costo</label>
                                <select id="f_tipoCosto" class="form-control"></select>
                            </div>
                        </form>
                    </div>
                </div>
            </div>
            <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>COSTOS LOGÍSTICOS ($)
                           
                            <span id="add"><i class="fa fa-plus"></i>NUEVO</span>
                        </h3>
                    </div>
                    <div class="panel-body">
                        <table id="grid"></table>
                        <div id="pager"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="add-dialog" class="dialog">
        <form id="add-form">
            <div class="form-group">
                <label>Ciudad</label>
                <select id="add-ciudad" class="form-control" required></select>
            </div>
            <div class="form-group">
                <label>Tipo Costo</label>
                <div class="input-group">
                    <select id="add-tipo-costo" class="form-control" required></select>
                    <span class="input-group-addon" id="add-type" title="Agregar Tipo de Costo">
                        <i class="fa fa-plus"></i>
                    </span>
                </div>
                
            </div>
            <div class="form-group">
                <label>Valor</label>
                <input type="number" value="0" min="0" class="form-control" id="add-value" required />
            </div>
        </form>

    </div>

    <div id="edit-dialog" class="dialog">
        <form id="edit-form">
            <div class="form-group">
                <label>Ciudad</label>
                <select id="edit-ciudad" class="form-control" disabled></select>
            </div>
            <div class="form-group">
                <label>Tipo Costo</label>
                <select id="edit-tipo-costo" class="form-control" disabled></select>
            </div>
            <div class="form-group">
                <label>Valor</label>
                <input type="number" value="0" min="0" class="form-control" id="edit-value" required />
            </div>
        </form>

    </div>
    <div id="add-type-dialog" class="dialog">
        <form id="add-type-form">
            <div class="form-group">
                <label>Nombre</label>
                <input type="text" class="form-control" placeholder="Nombre" id="add-type-name" required />
            </div>
        </form>
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
<script src="bower_components/linq.min.js"></script>
<script src="js/menu.js?20062017"></script>
<script src="js/jquery.Rut.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/costoslogisticos.js?20062017"></script>
</html>
