
var __filas = 15;

loadDataConfig();

//function listaritemizados() {
//    loadDataIt();
//}

function loadDataConfig() {
    var totregistros = 0;
    var paginas = 0;

    //var _busqueda = ""; //$("#ctrl_fec_ini").val();
    //var _activo = $("#activo").val();
    //var _origen = $("#busca_origen").val();
    //var _numreq = $("#busca_numreq").val();
    //var _origen = $("#busca_origen").val();
    var _sw = 0;

    $.ajax({
        url: "/Configuracion/List",
        type: "GET",
        data: {
            busqueda: '',
            estado: 1,
            page: 1,
            filas: __filas
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            //console.log(result);

            $('#tabla_objeto_par').html("");
            $.each(JSON.parse(result), function (key, item) {
                totregistros = item.filas;
                _sw = 1; /// para verificar que tiene registros

                armarLista(item.item, item.id_parametro, item.descripcion_parametro, item.valor_parametro, item.estado_parametro);

            });

            ///////////////////////////////////////////////////
            /////////////////PAGINADO
            ///////////////////////////////////////////////////
            paginas = Math.floor(totregistros / __filas);
            if (totregistros % __filas > 0) {
                paginas = paginas + 1;
            }

            $('#paginado_par').html("");
            if (paginas > 1) {
                for (var i = 1; i <= paginas; i++) {

                    $('#paginado_par')
                        .append($('<li class="page-item" />').html('<a class="page-link" href="#" onclick="listarpagina(\'' + '' + '\',' + i + ',' + __filas + ')">' + i + '</a>'));

                }

            } else {
                if (_sw == 0) {
                    $('#tabla_objeto_par')
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
    //alert('pagina ' + _pagina + '  filas ' + _filas + '  reque ' + _numreq);


    $.ajax({
        url: "/Configuracion/List",
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

            $('#tabla_objeto_par').html("");
            $.each(JSON.parse(result), function (key, item) {

                armarLista(item.item, item.id_parametro, item.descripcion_parametro, item.valor_parametro, item.estado_parametro);

            });

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}


function armarLista(_item, _id, _descripcion, _valor, _estado) {

    var _opt = '';
    //console.log(" estado : " + _estado);
    //alert(parseInt(_estado));

    if (_estado == 1) {
        //alert("Verdadero");

        var _opt_leer = '<a href="#" class="dropdown-item" onclick="verDatosItem(' + _id + ')"><i class="fas fa-eye"></i>&nbsp;&nbsp;&nbsp;ACTUALIZAR</a>';
        //var _opt_inac = '<a href="#" class="dropdown-item" onclick="InactivarItemi(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-window-close"></i>&nbsp;&nbsp;&nbsp;CERRAR</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default"><i class="fas fa-cog"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        _opt = _opt + '        ' + _opt_leer;
        //_opt = _opt + '        ' + _opt_inac;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    var _color = "white";
    $('#tabla_objeto_par')
        .append(
            $('<tr style="background-color:' + _color + ';"/>')
                .append($('<td/>').html(_opt))
                .append($('<td/>').html(_descripcion))
                .append($('<td/>').html(_valor))
        );

}


//function Eliminar(_origen, _idvisita, _item, _foto) {
//    swal({
//        title: "¿Seguro(a) de eliminar revisión?",
//        text: "",
//        icon: "warning",
//        buttons: true,
//        dangerMode: true,
//        buttons: ['CANCELAR', 'SI, SEGURO']
//    })
//        .then((willDelete) => {
//            if (willDelete) {
//                //swal("Solicitud de tramite recepcionada", {
//                //    icon: "success",
//                //});               
//                elimina_revision(_origen, _idvisita, _item, _foto);
//            } else {
//                swal("Operación cancelada");
//            }
//        });
//}


//function elimina_revision(_origen, _idvisita, _item, _foto) {

//    $.ajax({
//        url: "/Revision/Eliminararchivo",
//        type: "POST",
//        data: {
//            origen: _origen,
//            idvisita: _idvisita,
//            item: _item,
//            foto: _foto
//        },
//        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
//        success: function (data) {

//            //console.log(data);

//            if (data.operacion == true) {
//                toastr.success(data.mensaje);
//                loadDataRev();
//                //$("#modal-activo").modal('hide');
//            }
//            else {
//                toastr.error(data.mensaje);
//            }

//            if (data.operacion2 == false) {
//                toastr.error(data.mensaje2);
//            }

//        },
//        error: function () {
//            toastr.error("Ocurrio un error!!");
//        }
//    });
//}


function cambiar() {
    var _tipo = $('#v_tipo').val();
    var _div_producto = document.getElementById('m_codigo');

    var _uni = document.getElementById('v_chrunidad');
    var _glosa = document.getElementById('v_glosa');

    var _glosatxt = document.getElementById('GLOSA1');
    var _glosaarea = document.getElementById('GLOSA2');

    if (_tipo == "001") {
        //alert("001");
        _div_producto.style.display = 'block';
        $('#v_chrunidad').val("0");
        $('#v_chrunidad').change();
        //_uni.disabled = true;
        //_glosa.disabled = true;

        _glosaarea.style.display = 'none';
        _glosatxt.style.display = 'block';
    }

    if (_tipo == "002") {
        //alert("002");
        _div_producto.style.display = 'none';

        $('#v_chrunidad').val("01");
        $('#v_chrunidad').change();
        _uni.disabled = true;

        _glosaarea.style.display = 'block';
        _glosatxt.style.display = 'none';
    }

    if (_tipo == "003") {
        //alert("003");
        _div_producto.style.display = 'none';
        $('#v_chrunidad').val("01");
        $('#v_chrunidad').change();
        _uni.disabled = true;
        _glosa.disabled = false;

        _glosaarea.style.display = 'none';
        _glosatxt.style.display = 'block';
    }

    if (_tipo == "004") {
        //alert("003");
        _div_producto.style.display = 'none';
        $('#v_chrunidad').val("01");
        $('#v_chrunidad').change();
        _uni.disabled = true;
        _glosa.disabled = false;

        _glosaarea.style.display = 'none';
        _glosatxt.style.display = 'block';
    }

}

async function myLeerVisita(_id) {

    const result = await $.ajax({
        url: "/Visita/getData",
        type: "POST",
        data: {
            id: _id
        },
        contentType: 'application/x-www-form-urlencoded',
        datatype: "html",
        //success: function (result) {
        //success: function (result) {

        //},
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

    return result;

}

async function limpiarFormularioPeriodo(/*_origen, _numreq*/) {

    //var _id = $('#busca_intvisita2').val();
    //const data = await myLeerVisita(_id);
    //var _sw = 1;
    //$.each(JSON.parse(data), function (key, item) {
    //    //alert(item.nombre_solicitante);
    //    if (item.ESTADOREQ == '2') {
    //        _sw = 0;
    //        alert_error("Imposible registrar", "Requerimiento cambio de estado");
    //    }

    //});

    //if (_sw == 1) {
    //    loadCombo(0, 'Itemizado', 'ListUnidades', 'v_chrunidad', 1);
    //    $("#v_id").val("0");
    //    $("#v_numreq").val(_numreq);
    //    $("#v_origen").val(_origen);

    //    $("#v_tipo").val("001");
    //    $("#v_tipo").change();

    //    var _div_codigo = document.getElementById('m_codigo');
    //    _div_codigo.style.display = 'block';
    //    $("#v_codigo").val("");
    //    $("#v_chrcodproducto").val("");

    //    $("#v_glosa").val("");
    //    $("#v_glosa2").val("");
    //    $("#v_glosa_2").val("");
    //    $("#v_cantidad").val("1");

    //    $("#v_grupo").val("");
    //    //$("#v_orden").val(item.INTORDEN);

    //    $("#v_criticidad").val("B");
    //    $("#v_criticidad").change();


    //    var _div_orden = document.getElementById('m_orden');
    //    _div_orden.style.display = 'none';

    //    var _uni = document.getElementById('v_chrunidad');
    //    var _glosa = document.getElementById('v_glosa');
    //    _uni.disabled = false;
    //    _glosa.disabled = false;

    //    $('#modal-sm-itemizado').modal('show');
    //}

    $('#modal-sm-periodo').modal('show');

}


function GuardarPar() {

    if ($.trim($("#v_descripcion_parametro").val()) == '') {
        toastr.error("Ingrese descripción");
        return false;
    }

    if ($.trim($("#v_valor_parametro").val()) == '') {
        toastr.error("Ingrese valor");
        return false;
    }

    /*if (id == "0") {*/
    grabaitem('Actualizar');
    //} else {
    //    grabaitem('Guardar');
    //}

}


function grabaitem(_accion) {

    var formData = $('#frmParametroRegistrar').serialize();

    $.ajax({
        url: "/Configuracion/" + _accion,
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataConfig();
                $("#modal-sm-parametro").modal('hide');
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

async function verDatosItem(_id) {
    
        //var _id = $('#busca_intvisita2').val();

        $.ajax({
            type: "POST",
            url: "/Configuracion/getData",
            data: {
                id: _id
            },
            datatype: "html",
            contentType: 'application/x-www-form-urlencoded',
            success: function (data) {
                //alert("retorna data");
                //console.log("JIM");
                //console.log(data);
                $.each(JSON.parse(data), function (key, item) {

                    $("#v_id_parametro").val(item.id_parametro);
                    $("#v_descripcion_parametro").val(item.descripcion_parametro);
                    $("#v_valor_parametro").val(item.valor_parametro);

                });

            },
            error: function () {
                toastr.error("Ocurrio un error!!");
            }
        });


        $('#modal-sm-parametro').modal('show');
    

}


async function InactivarItemi(_anio, _mes, _nombre_mes) {

    //var _id = $('#busca_intvisita2').val();
    //const data = await myLeerVisita(_id);
    //var _sw = 1;
    //$.each(JSON.parse(data), function (key, item) {
    //    //alert(item.nombre_solicitante);
    //    if (item.ESTADOREQ == '2') {
    //        _sw = 0;
    //        alert_error("Imposible inactivar", "Requerimiento cambio de estado");
    //    }

    //});

    //if (_sw == 1) {
    swal({
        title: "¿Seguro(a) de cerrar el periodo?",
        text: _anio + "-" + _nombre_mes,
        icon: "warning",
        buttons: true,
        dangerMode: true,
        buttons: ['CANCELAR', 'SI, SEGURO']
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
        title: "¿Seguro(a) de abrir el periodo?",
        text: _anio + "-" + _nombre_mes,
        icon: "warning",
        buttons: true,
        dangerMode: true,
        buttons: ['CANCELAR', 'SI, SEGURO']
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
    //alert("inactivo");
    //var _id = $("#v_idina").val();
    //alert("activar " + _id);

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
                toastr.error(data.mensaje);
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

function buscarproducto() {
    var _cadena = $("#v_busqueda").val();

    if ($.trim(_cadena) == "") {
        alert_error("Ingrese texto", "");
        return false;
    }

    $.ajax({
        url: "/Itemizado/ListProductos",
        type: "GET",
        data: {
            cadena: _cadena
        },
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (result) {

            //console.log(data);
            $('#tabla_productos').html("");
            $.each(JSON.parse(result), function (key, item) {

                //armarLista(item.item, item.CHRCODIGOORIGEN, item.INTITEM, item.CHRCODIGOPRODUCTO, item.vchdescproducto, item.VCHGLOSA, item.VCHGRUPO, item.NUMCANTIDAD, item.CHRUNIDAD, item.unidad, item.CHRCRITICIDAD, item.criticidad, item.INTORDEN, item.CHRTIPO, item.tipo, _activo, _numreq);
                $('#tabla_productos')
                    .append(
                        $('<tr/>')
                            //// ,\'' + item.VCHCODIGOPRODUCTO + '\',\'' + '-' + item.VCHDESCPRODUCTO + '-' + '\'
                            .append($('<td/>').html('<button type="button" title="Agregar ' + item.VCHCODIGOPRODUCTO + '" onclick="agregarproducto(\'' + item.CHRCODIGOPRODUCTO + '\',\'' + item.UNIDAD + '\');" class="btn btn-secondary">&nbsp;&nbsp;+&nbsp;&nbsp;</button>'))
                            .append($('<td/>').html(item.VCHCODIGOPRODUCTO))
                            .append($('<td/>').html(item.VCHDESCPRODUCTO))
                            .append($('<td/>').html(item.CHRCODIGOFAB))
                            .append($('<td/>').html(item.VCHDESCMARCA))

                        //.append($('<td  class= "text-right py-0 align-middle"/>').html('<div class="btn-group btn-group-sm"><a href="#" class="btn btn-primary" data-toggle="modal" data-target="#modal-sm" onclick="verDatos()"><i class="fas fa-eye"></i></a><a href="#" class="btn btn-danger" data-toggle="modal" data-target="#modal-activo" onclick="ModalActivo()"><i class="fas fa-trash"></i></a></div>'))
                        //.append($('<td  class= "text-right py-0 align-middle"/>').html(_opt_fin))
                    );
                //alert(item.VCHCODIGOPRODUCTO);

            });

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
        }
    });

}


function agregarproducto(_codprod/*, _vchcodprod, _descprod*/, _unidad) {

    $.ajax({
        url: "/Itemizado/ObtenerProducto",
        type: "GET",
        data: {
            codigoproducto: _codprod
        },
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (result) {

            //console.log(data);
            //$('#tabla_productos').html("");
            $.each(JSON.parse(result), function (key, item) {

                $('#v_codigo').val(item.VCHCODIGOPRODUCTO);
                $('#v_chrcodproducto').val(_codprod);
                $('#v_glosa').val(item.VCHDESCPRODUCTO);
                $('#v_glosa_2').val(item.VCHDESCPRODUCTO);

                $("#v_chrunidad").val(_unidad);
                $("#v_chrunidad").change();

                $("#modal-sm-busca-prod").modal('hide');
                $('#v_busqueda').val("");
                $('#tabla_productos').html("");

                var _uni = document.getElementById('v_chrunidad');
                var _glosa = document.getElementById('v_glosa');
                _uni.disabled = true;
                _glosa.disabled = true;


            });

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
        }
    });


}


function impresionItemizado(_origen, _numreq, _asunto) {
    //alert(_INTVISITA);
    $.ajax({
        type: "GET",
        cache: false,
        url: "/Itemizado/Impresion",
        data: {
            origen: _origen,
            numreq: _numreq,
            asunto: _asunto
        },
        contentType: true, // Not to set any content header  
        processData: true, // Not to process data  
        success: function (data) {

            //    if (data.operacion == true) {

            //        load_files(_idot);
            //        alert_success(data.mensaje, "");

            //    } else {
            //        alert_error(data.mensaje, '');
            //    }
            window.location = '/Itemizado/Download?fileGuid=' + data.FileGuid
                + '&filename=' + data.FileName;

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
        }
    });
}


function eventoenter(e) {
    /*alert("enter");*/
    //tecla = (document.all) ? e.keyCode : e.which;
    //return (tecla != 13);
    if (e.which == 13) {
        //alert("enter");
        //alert("enter");
        buscarproducto();
        return false;
    }
}


async function inicializarCopy(_origen, _numreq) {
    $('#v_origen2').val(_origen);
    $('#v_num_req_destino').val(_numreq);
    loadComboempresa(0, 'Itemizado', 'ListVisitasCopy', 'd_num_visita', 1, _origen);
    $('#modal-sm-copy').modal('show');
}


async function inicializarExcel(_origen, _numreq) {
    //$('#v_origen2').val(_origen);
    //$('#v_num_req_destino').val(_numreq);
    //loadComboempresa(0, 'Itemizado', 'ListVisitasCopy', 'd_num_visita', 1, _origen);
    //$('#modal-sm-copy').modal('show');

    //var _busqueda = $("#busqueda").val();
    //var _activo = 1; //$("#activo").val();
    //var _fecha1 = $("#fechaini").val();
    //var _fecha2 = $("#fechafin").val();

    var _activo = $("#activo").val();
    //var _origen = $("#busca_origen").val();
    //var _numreq = $("#busca_numreq").val();


    $.ajax({
        cache: false,
        url: "/Itemizado/Exportar_lista_exp",
        data: {
            origen: _origen,
            numreq: _numreq,
            activo: _activo
        },
        dataType: "json",
        success: function (data) {

            window.location = '/Itemizado/DownloadExcel?fileGuid=' + data.FileGuid
                + '&filename=' + data.FileName;

        },
        error: function (errormessage) {
            toastr.error(errormessage.responseText);
        }
    })

}


function GuardarCopy() {
    var _copy = $('#d_num_visita').val();

    if ($.trim(_copy) == '0') {
        toastr.error("Seleccione requerimiento");
        return false;
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////
    //////////////////////////////////////////////////////////////////////////////////////////////////////////

    var formData = $('#frmItemizadoCopiar').serialize();
    //console.log(formData);

    $.ajax({
        url: "/Itemizado/GuardarCopy",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataIt();
                $("#modal-sm-copy").modal('hide');
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