<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Normas.aspx.cs" Inherits="Normas" %>

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
    <link href="css/certel.css" rel="stylesheet" />
    <link href="css/elementos/normas.css" rel="stylesheet"/>
</head>
<body>
    <div id="cssmenu"></div>
    <div class="container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="page-header">
                    <h1>Normas</h1>
                </div>
            </div>
        </div>
        <div class="row">
             <div class="col-xs-12 col-sm-12 col-md-3 col-lg-3">
                 <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>FILTROS <span><i class="fa fa-remove"></i></span></h3>
                    </div>
                    <div class="panel-body" data-height-full="true">
                        <form id="formFiltros">
                            <div class="form-group">
                                <label for="f_nombre">Nombre Norma</label>
                                <input type="text" class="form-control" id="f_nombre" name="nombre"/>
                            </div>
                        </form>
                    </div>
                </div>
             </div>
            <div class="col-xs-12 col-sm-12 col-md-9 col-lg-9">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3>REGISTROS <span id="add"><i class="fa fa-plus"></i></span><span id="reload"><i class="fa fa-refresh"></i></span></h3>
                    </div>
                    <div class="panel-body" data-height-full="true" id="pb">
                        <div id="panel">
                            <div class="col-xs-6 panel-flex" id="panelNormas">

                            </div>
                            <div class="col-xs-6 panel-flex" id="panelNormas2"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <%--Foot-Page--%>
    <div class="foot-page">Certel - Certificación en Elevación S.A.</div>

    <%--Lock Screen--%>
    <div class="lock-screen">
        <i class="fa fa-cog fa-spin fa-3x fa-fw margin-bottom"></i>
        <span class="sr-only">Cargando...</span>
    </div>
    <%--Dialogs--%>
    <%--Add--%>
    <div id="add-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-dialog-form">
                        <div class="form-group">
                            <label for="add-nombre">Nombre (*)</label>
                            <input type="text" placeholder="Ej. NCh3395:2016" class="form-control" id="add-nombre" maxlength="50" required />
                        </div>
                        <div class="form-group">
                            <label for="add-tipo">Tipo (*)</label>
                            <select id="add-tipo" class="form-control" required>
                                <option></option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="add-titulo-regulacion">Título Regulación (*)</label>
                            <input type="text" placeholder="Ej. Requisitos para la inspección de ascensores y montacargas eléctricos existentes." class="form-control" id="add-titulo-regulacion" maxlength="100" required/>
                        </div>
                        
                        <div class="form-group" id="a_fg-tipo-informe">
                            <label for="add-tipo-informe">Tipo Informe</label>
                            <select id="add-tipo-informe" class="form-control">
                                <option value="0">Seleccione...</option>
                            </select>
                        </div>
                        <div class="form-group" id="a_fg-parrafo" style="display:none">
                            <label for="add-parrafo">Párrafo Introductorio</label>
                            <textarea class="form-control" id="add-parrafo" maxlength="5000"></textarea>
                        </div>
                        <div class="required-data-label">(*) = Campo Obligatorio</div>
                    </form>
                </div>
            </div>
        </div>    
    </div>

    <%--Edit--%>
    <div id="edit-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <div id="tabs">
                        <ul>
                            <li><a href="#edit">Editar</a></li>
                            <li><a href="#requisitos">Requisitos</a></li>
                            <li><a href="#terminos">Términos y definiciones</a></li>
                        </ul>
                        <div id="edit">
                            <form id="edit-dialog-form">
                                <div class="form-group">
                                    <label for="edit-nombre">Nombre (*)</label>
                                    <input type="text" placeholder="Ej. NCh3395:2016" class="form-control" id="edit-nombre" maxlength="50" required />
                                </div>
                                <div class="form-group">
                                    <label for="edit-tipo">Tipo (*)</label>
                                    <select id="edit-tipo" class="form-control" required disabled>
                                        <option></option>
                                    </select>
                                </div>
                                <div class="form-group">
                                    <label for="edit-titulo-regulacion">Título Regulación (*)</label>
                                    <input type="text" placeholder="Ej. Requisitos para la inspección de ascensores y montacargas eléctricos existentes." class="form-control" id="edit-titulo-regulacion" maxlength="100" required />
                                </div>
                                
                                <div class="form-group" id="fg-tipo-informe">
                                    <label for="edit-tipo-informe">Tipo Informe</label>
                                    <select id="edit-tipo-informe" class="form-control">
                                        <option value="0">Seleccione...</option>
                                    </select>
                                </div>
                                <div class="form-group" id="fg-parrafo" style="display:none">
                                    <label for="edit-parrafo">Párrafo Introductorio</label>
                                    <textarea class="form-control" id="edit-parrafo" maxlength="5000"></textarea>
                                </div>
                                <div class="required-data-label">(*) = Campo Obligatorio</div>
                            </form>
                        </div>
                        <div id="requisitos">
                            <button class="btn btn-default" id="add-title">
                                <i class="fa fa-plus"></i>
                                NUEVO TÍTULO
                            </button>
                            <div id="panel-requisitos"></div>
                        </div>
                        <div id="terminos">
                            <ul id="terminos-list" class="list-group"></ul>
                            <form id="terminos-form">
                                <div class="form-group">
                                    <label>Término</label>
                                    <input type="text" id="termino" placeholder="Término" class="form-control" required/>
                                </div>
                                <div class="form-group">
                                    <label>Definición</label>
                                    <div class="input-group">
                                        <textarea name="definicion" id="definicion" class="form-control" placeholder="Definición" required></textarea>
                                        <span class="input-group-addon btn btn-default" id="terminos-submit">
                                            <i class="fa fa-save"></i> Guardar
                                        </span>
                                    </div>
                                    
                                </div>
                                
                            </form>
                        </div>
                    </div>
                    
                </div>
            </div>
        </div>    
    </div>
    <%--New Title--%>
    <div id="add-title-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-title-dialog-form">
                        <div class="form-group">
                            <label for="add-title-name">Título (*)</label>
                            <input type="text" maxlength="500" class="form-control" id="add-title-title" required/>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--Edit Title--%>
    <div id="edit-title-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="edit-title-dialog-form">
                        <div class="form-group">
                            <label for="add-title-name">Título (*)</label>
                            <input type="text" maxlength="500" class="form-control" id="edit-title-title" required/>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--Grilla Requisitos--%>
    <div id="grid-requisitos-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-6">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <h3>REQUISITOS <span id="add-requisito"><i class="fa fa-plus"></i></span></h3>
                        </div>
                        <div class="panel-body">
                            <table id="grid-requisitos"></table>
                            <div id="pager-grid-requisitos"></div>
                        </div>
                    </div>
                    
                </div>
                <div class="col-xs-6">
                    <div class="panel panel-default" id="panel-caracteristicas">
                        <div class="panel-heading">
                            <h3>CARACTERÍSTICAS <span id="add-caracteristica"><i class="fa fa-plus"></i></span></h3>
                        </div>
                        <div class="panel-body">
                             <table id="grid-caracteristicas"></table>
                            <div id="pager-grid-caracteristicas"></div>
                        </div>
                    </div>
                   
                </div>
            </div>
        </div>
    </div>

    <%--New Requisito--%>
    <div id="add-requisito-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-requisito-dialog-form">
                        <div class="form-group">
                            <label for="add-requisito-text">Texto (*)</label>
                            <textarea class="form-control" maxlength="200" placeholder="Texto requisito" id="add-requisito-text" required></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--New Caracteristica--%>
    <div id="add-caracteristica-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="add-caracteristica-dialog-form">
                        <div class="form-group">
                            <label for="add-requisito-text">Texto (*)</label>
                            <textarea class="form-control" maxlength="5000" placeholder="Texto característica" rows="10" id="add-caracteristica-text" required></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--Edit Requisito--%>
    <div id="edit-requisito-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="edit-requisito-dialog-form">
                        <div class="form-group">
                            <label for="edit-requisito-text">Texto (*)</label>
                            <textarea class="form-control" maxlength="200" placeholder="Texto requisito" id="edit-requisito-text" required></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <%--Edit Caracteristica--%>
    <div id="edit-caracteristica-dialog" class="dialog">
        <div class="container-fluid">
            <div class="row">
                <div class="col-xs-12">
                    <form id="edit-caracteristica-dialog-form">
                        <div class="form-group">
                            <label for="edit-caracteristica-text">Texto (*)</label>
                            <textarea class="form-control" maxlength="5000" placeholder="Texto característica" rows="10" id="edit-caracteristica-text" required></textarea>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <div class="dialog" id="edit-terminos">
        <form id="edit-terminos-form">
            <div class="form-group">
                <label>Término</label>
                <input type="text" id="edit-termino" placeholder="Término" class="form-control" required />
            </div>
            <div class="form-group">
                <label>Definición</label>
                <div class="input-group">
                    <textarea name="definicion" id="edit-definicion" class="form-control" placeholder="Definición" required></textarea>
                    <span class="input-group-addon btn btn-default" id="edit-terminos-submit">
                        <i class="fa fa-save"></i> Guardar
                    </span>
                </div>
                                    
            </div>
                                
        </form>
    </div>
</body>
<script src="bower_components/jquery/dist/jquery.min.js"></script>
<script src="bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
<script src="bower_components/jquery-ui/jquery-ui.min.js"></script>
<script src="bower_components/jqGrid/jquery.jqGrid.min.js"></script>
<script src="bower_components/jqGrid/grid.locale-es.js"></script>
<script src="bower_components/alertify.js/alertify.min.js"></script>
<script src="bower_components/select2/dist/js/select2.min.js"></script>
<script src="js/menu.js"></script>
<script src="js/certel.js"></script>
<script src="js/elementos/normas.js"></script>
</html>
