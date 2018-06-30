<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ListasDePrecio.aspx.cs" Inherits="ListasDePrecio" %>

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
    <link href="css/certel.css" rel="stylesheet" />
</head>
<body>
    <div id="cssmenu"></div>
    <div class="col-xs-12 col-sm-12 col-md-3 col-lg-3">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3>filtros <span id="f_remove"><i class="fa fa-times"></i></span></h3>
            </div>
            <div class="panel-body">
                <form id="filter-form">
                    <div class="form-group">
                        <label>Elevador</label>
                        <select id="f_elevador" class="form-control"></select>
                    </div>
                    <div class="form-group">
                        <label>Paradas</label>
                        <input id="f_paradas" type="number" value="0" min="0" class="form-control" />
                    </div>
                    <div class="form-group">
                        <label>Más de 1 equipo</label>
                        <input type="checkbox" id="f_moreThanOne" checked />
                    </div>
                    <div class="form-group">
                        <label>Producto</label>
                        <select id="f_producto" class="form-control"></select>
                    </div>
                    <div class="form-group">
                        <button type="submit" class="btn btn-default btn-block">
                            <i class="fa fa-search"></i> Buscar
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3>Listas de precio (UF)
                 <span id="add-lista"><i class="fa fa-plus"></i></span>
                </h3>
            </div>
            <div class="panel-body">
                <div class="row">
                    <table id="grid"></table>
                    <div id="pager"></div>
                </div>

            </div>
        </div>
    </div>



    <div id="dialog-nueva-lista" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="form-nueva-lista">

                        <div class="form-group col-xs-12">
                            <label for="tipo-elevador">Tipo Elevador (*)</label>
                            <select class="form-control" id="n_tipo-elevador" required>
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="Paradas">Paradas (*)</label>
                            <input class="form-control" type="number" id="n_Paradas" required/>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="">¿Mas de un Equipo? (*)</label>
                            <input type="checkbox" id="n_mas-de-un-equipo" required/>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="add-ingeniero">Tipo Producto (*)</label>

                            <select class="form-control" id="n_tipo-Producto" required>
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="m_valor">Valor UF (*)</label>
                            <input class="form-control" type="number" id="n_valor" required />
                        </div>


                    </form>
                </div>
            </div>
        </div>
    </div>


    <div id="dialog-mod-lista" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="form-mod-lista">

                        <div class="form-group col-xs-12" style="display: none">
                            <input type="text" id="m_id" />
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="tipo-elevador">Tipo Elevador (*)</label>
                            <select class="form-control" id="m_tipo-elevador" required disabled>
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="Paradas">Paradas (*)</label>
                            <input class="form-control" type="number" id="m_Paradas" required disabled/>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="m_mas-de-un-equipo">¿Mas de un Equipo? (*)</label>
                            <input  type="checkbox" id="m_mas-de-un-equipo" required disabled/>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="add-ingeniero">Tipo Producto (*)</label>

                            <select class="form-control" id="m_tipo-Producto" required disabled>
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>

                        <div class="form-group col-xs-12">
                            <label for="m_valor">Valor UF (*)</label>
                            <input class="form-control" type="number" id="m_valor" required />
                        </div>


                    </form>
                </div>
            </div>
        </div>
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
<script src="js/menu.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/ListasDePrecio.js"></script>
</html>
