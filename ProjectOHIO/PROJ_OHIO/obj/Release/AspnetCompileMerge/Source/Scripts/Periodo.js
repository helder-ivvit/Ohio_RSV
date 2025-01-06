
var __filas = 15;

loadDataIt();

function loadDataIt() {
    var totregistros = 0;
    var paginas = 0;
    var _sw = 0;

    $.ajax({
        url: "/Periodo/List",
        type: "GET",
        data: {
            busqueda: $("#v_anio_buscar").val(),
            estado: 1,
            page: 1,
            filas: __filas            
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            //console.log(result);

            $('#tabla_objeto_per').html("");
            $.each(JSON.parse(result), function (key, item) {
                totregistros = item.filas;
                _sw = 1; /// para verificar que tiene registros
                
                armarLista(item.item, item.anio_periodo, item.mes_periodo, item.nombre_mes, item.estado_periodo, item.nombre_estado);

            });

            ///////////////////////////////////////////////////
            /////////////////PAGINADO
            ///////////////////////////////////////////////////
            paginas = Math.floor(totregistros / __filas);
            if (totregistros % __filas > 0) {
                paginas = paginas + 1;
            }

            $('#paginado_per').html("");
            if (paginas > 1) {
                for (var i = 1; i <= paginas; i++) {

                    $('#paginado_per')
                        .append($('<li class="page-item" />').html('<a class="page-link" href="#" onclick="listarpagina(\'' + '' + '\',' + i + ',' + __filas + ')">' + i + '</a>'));

                }

            } else {
                if (_sw == 0) {
                    $('#tabla_objeto_per')
                        .append(
                            $('<tr/>')
                                .append($('<td/>').html(""))
                                .append($('<td/>').html("No hay registros"))
                                .append($('<td/>').html(""))
                        );
                }
            }

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}


function listarpagina(_busqueda, _pagina, _filas) {

    $.ajax({
        url: "/Periodo/List",
        type: "GET",
        data: {
            busqueda: _busqueda,
            estado: 1,
            page: _pagina,
            filas: _filas
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            console.log(result);

            $('#tabla_objeto_per').html("");
            $.each(JSON.parse(result), function (key, item) {

                armarLista(item.item, item.anio_periodo, item.mes_periodo, item.nombre_mes, item.estado_periodo, item.nombre_estado);

            });

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}


function armarLista(_item, _anio, _mes, _nombre_mes, _estado, _nombre_estado) {

    var _opt = '';
    //console.log(" estado : " + _estado);
    //alert(parseInt(_estado));

    if (_estado == 1) {
        //alert("Verdadero");

        //var _opt_leer = '<a href="#" class="dropdown-item" onclick="verDatosItem(\'' + _chrcodigoorigen + '\',\'' + _numreq + '\',' + _INTITEM + ')"><i class="fas fa-eye"></i>&nbsp;&nbsp;&nbsp;ACTUALIZAR</a>';
        var _opt_inac = '<a href="#" class="dropdown-item" onclick="InactivarItemi(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-window-close"></i>&nbsp;&nbsp;&nbsp;CERRAR</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default"><i class="fas fa-cog"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        //_opt = _opt + '        ' + _opt_leer;
        _opt = _opt + '        ' + _opt_inac;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    if (_estado == 0) {
        //alert("falso");

        var _opt_act = '<a href="#" class="dropdown-item" onclick="ActivarItemi(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-key"></i>&nbsp;&nbsp;&nbsp;REAPERTURA</a>';
        //#D32F2F
        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default bg-danger"><i class="fas fa-unlock-alt"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon bg-danger" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        _opt = _opt + '        ' + _opt_act;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    var _color = "white";

    $('#tabla_objeto_per')
        .append(
            $('<tr style="background-color:' + _color + ';"/>')
                .append($('<td/>').html(_opt))
                .append($('<td/>').html(_anio))
                .append($('<td/>').html(_nombre_mes))
                .append($('<td/>').html(_nombre_estado))
            //.append($('<td  class= "text-right py-0 align-middle"/>').html('<div class="btn-group btn-group-sm"><a href="#" class="btn btn-primary" data-toggle="modal" data-target="#modal-sm" onclick="verDatos()"><i class="fas fa-eye"></i></a><a href="#" class="btn btn-danger" data-toggle="modal" data-target="#modal-activo" onclick="ModalActivo()"><i class="fas fa-trash"></i></a></div>'))
            //.append($('<td  class= "text-right py-0 align-middle"/>').html(_opt_fin))
        );

}


async function limpiarFormularioPeriodo() {
    
    $('#modal-sm-periodo').modal('show');

}


function GuardarIt() {
    
    if ($.trim($("#v_anio").val()) == '') {
        toastr.error("Ingrese año");
        return false;
    }

    if ($.trim($("#v_mes").val()) == '') {
        toastr.error("Seleccione mes");
        return false;
    }

    if ($.trim($("#v_mes").val()) == 0) {
        toastr.error("Seleccione mes");
        return false;
    }


    /*if (id == "0") {*/
        grabaitem('Guardar');
    //} else {
    //    grabaitem('Actualizar');
    //}

}


function grabaitem(_accion) {

    var formData = $('#frmPeriodoRegistrar').serialize();
    
    $.ajax({
        url: "/Periodo/" + _accion,
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataIt();
                $("#modal-sm-periodo").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            //alert("Ocurrio un error!!");
        }
    });

}



async function InactivarItemi(_anio, _mes, _nombre_mes) {

    //if (_sw == 1) {
        swal({
            title: "¿Seguro de cerrar el periodo " + _nombre_mes + "-" + _anio +"?",
            text: "",
            icon: "warning",
            buttons: true,
            dangerMode: true,
            buttons: ['CANCELAR', 'ACEPTAR']
        })
            .then((willDelete) => {
                if (willDelete) {
                    //swal("Solicitud de tramite recepcionada", {
                    //    icon: "success",
                    //});               
                    cambiarestado(_anio, _mes, 'Activo');
                } else {
                    swal("Operación cancelada");
                }
            });
    //}

}


function ActivarItemi(_anio, _mes, _nombre_mes) {
    swal({
        title: "Periodo " + _nombre_mes + "-" + _anio + " se encuentra cerrado. ¿Desea reaperturar el periodo?",
        text: "",
        icon: "warning",
        buttons: true,
        dangerMode: true,
        buttons: ['CANCELAR', 'ACEPTAR']
    })
        .then((willDelete) => {
            if (willDelete) {
                //swal("Solicitud de tramite recepcionada", {
                //    icon: "success",
                //});               
                cambiarestado(_anio, _mes, 'Inactivo');
            } else {
                swal("Operación cancelada");
            }
        });
}


function cambiarestado(_anio, _mes, _accion) {

    $.ajax({
        url: "/Periodo/" + _accion,
        type: "POST",
        data: {
            anio: _anio,
            mes: _mes
        },
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);
                loadDataIt();
                //$("#modal-activo").modal('hide');
            }
            else {
                //toastr.error(data.mensaje);
                alert_error(data.mensaje);
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
        }
    });
}

