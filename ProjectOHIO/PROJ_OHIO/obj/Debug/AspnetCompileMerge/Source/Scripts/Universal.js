
var __filas = 15;

loadDataUniversal();

//function listaritemizados() {
//    loadDataIt();
//}

function loadDataUniversal() {
    var totregistros = 0;
    var paginas = 0;

    var _busqueda = "";
    var _sw = 0;

    $.ajax({
        url: "/Universal/List",
        type: "GET",
        data: {
            busqueda: _busqueda,
            estado: 1,
            page: 1,
            filas: __filas
        },
        contentType: "application/json;charset=utf-8",
        dataType: "json",
        success: function (result) {
            //console.log(result);

            $('#tabla_objeto_uni').html("");
            $.each(JSON.parse(result), function (key, item) {
                totregistros = item.filas;
                _sw = 1; /// para verificar que tiene registros

                //alert(item.item + '  ' + item.VCHFOTO);
                armarLista(item.item, item.anio_vida_universal, item.mes_vida_universal, item.nombre_mes, item.FlagDataPolizasSoles, item.FlagDataPolizasDolares, item.FlagDataCobranzasSoles, item.FlagDataCobranzasDolares, item.FlagDataEndoso, item.estado_periodo);

            });

            ///////////////////////////////////////////////////
            /////////////////PAGINADO
            ///////////////////////////////////////////////////
            paginas = Math.floor(totregistros / __filas);
            if (totregistros % __filas > 0) {
                paginas = paginas + 1;
            }

            $('#paginado_uni').html("");
            if (paginas > 1) {
                for (var i = 1; i <= paginas; i++) {

                    $('#paginado_uni')
                        .append($('<li class="page-item" />').html('<a class="page-link" href="#" onclick="listarpagina(\'' + _busqueda + '\',' + i + ',' + __filas + ')">' + i + '</a>'));

                }

            } else {
                if (_sw == 0) {
                    $('#tabla_objeto_uni')
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
        url: "/Universal/List",
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

            $('#tabla_objeto_uni').html("");
            $.each(JSON.parse(result), function (key, item) {

                armarLista(item.item, item.anio_vida_universal, item.mes_vida_universal, item.nombre_mes, item.FlagDataPolizasSoles, item.FlagDataPolizasDolares, item.FlagDataCobranzasSoles, item.FlagDataCobranzasDolares, item.FlagDataEndoso, item.estado_periodo);

            });

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}

function armarLista(_item, _anio, _mes, _nombre_mes, _FlagDataPolizasSoles, _FlagDataPolizasDolares,
    _FlagDataCobranzasSoles, _FlagDataCobranzasDolares, _FlagDataEndoso, _activo)
{

    var var__FlagDataPolizasSoles = '';
    var var__FlagDataPolizasDolares = '';
    var var__FlagDataCobranzasSoles = '';
    var var__FlagDataCobranzasDolares = '';
    var var__FlagDataEndoso = '';

    
    if (_FlagDataPolizasSoles == 0)
        var__FlagDataPolizasSoles = 'No';
    else
        var__FlagDataPolizasSoles = 'Si';

    if (_FlagDataPolizasDolares == 0)
        var__FlagDataPolizasDolares = 'No';
    else
        var__FlagDataPolizasDolares = 'Si';

    if (_FlagDataCobranzasSoles == 0)
        var__FlagDataCobranzasSoles = 'No';
    else
        var__FlagDataCobranzasSoles = 'Si';

    if (_FlagDataCobranzasDolares == 0)
        var__FlagDataCobranzasDolares = 'No';
    else
        var__FlagDataCobranzasDolares = 'Si';

    if (_FlagDataEndoso == 0)
        var__FlagDataEndoso = 'No';
    else
        var__FlagDataEndoso = 'Si';


    var _opt = '';
    var _anio_html = '';
    var _mes_html='';

    if (_activo == 1) {
        _anio_html = '<font color="#000000">' + _anio + '</font>';
        _mes_html = '<font color="#000000">' + _nombre_mes + '</font>';

        var _opt_polizas = '<a href="#" class="dropdown-item" onclick="mostrar_form_polizas(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de P&oacute;lizas</a>';
        var _opt_cobranzas = '<a href="#" class="dropdown-item" onclick="mostrar_form_cobranzas()"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Cobranzas</a>';
        var _opt_endosos = '<a href="#" class="dropdown-item" onclick="mostrar_form_endosos()"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Endosos</a>';
        var _opt_valor_poliza = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Valor P&oacute;liza</a>';
        var _opt_vida_univ = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Reserva Matematica VU</a>';
        var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
        var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matematica VU</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default"><i class="fas fa-cog"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        _opt = _opt + '        ' + _opt_polizas;
        _opt = _opt + '        ' + _opt_cobranzas;
        _opt = _opt + '        ' + _opt_endosos;
        _opt = _opt + '        ' + _opt_valor_poliza;
        _opt = _opt + '        ' + _opt_vida_univ;
        _opt = _opt + '        ' + _opt_result_valor_poliza;
        _opt = _opt + '        ' + _opt_result_vida_univ;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    if (_activo == 0) {
        // #c0392b
        _anio_html = '<font color="#ffa000">' + _anio + '</font>';
        _mes_html = '<font color="#ffa000">' + _nombre_mes + '</font>';

        var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
        var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matematica VU</a>';

        _opt = _opt + '<div class="btn-group">';
        _opt = _opt + '    <button type="button" class="btn btn-default bg-danger"><i class="fas fa-unlock-alt"></i></button>';
        _opt = _opt + '    <button type="button" class="btn btn-default dropdown-toggle dropdown-icon bg-danger" data-toggle="dropdown">';
        _opt = _opt + '        <span class="sr-only">Toggle Dropdown</span>';
        _opt = _opt + '    </button>';
        _opt = _opt + '    <div class="dropdown-menu" role="menu">';
        _opt = _opt + '        ' + _opt_result_valor_poliza;
        _opt = _opt + '        ' + _opt_result_vida_univ;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    var _color = "";
    


    $('#tabla_objeto_uni')
        .append(
            $('<tr style="background-color:' + _color + ';"/>')
                .append($('<td/>').html(_opt))
                .append($('<td/>').html(_anio_html))
                .append($('<td/>').html(_mes_html))
                .append($('<td/>').html(var__FlagDataPolizasSoles))
                .append($('<td/>').html(var__FlagDataPolizasDolares))
                .append($('<td/>').html(var__FlagDataCobranzasSoles))
                .append($('<td/>').html(var__FlagDataCobranzasDolares))
                .append($('<td/>').html(var__FlagDataEndoso))            
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

async function limpiarFormularioUniversal(/*_origen, _numreq*/) {

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

    //alert("001");
    loadCombo(0, 'Universal', 'ListPeriodos', 'v_periodo_universal', 1);
    //alert("002");
    $('#modal-sm-universal').modal('show');
    //alert("003");

}


function mostrar_form_polizas(_anio, _mes, _mesnombre) {

    $("#titulo_modal_poliza").html('Cargar Data de P&oacute;lizas <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(Periodo ' + _anio + '-' + _mesnombre + ')</b>');

    $("#v_anio_polizas").val(_anio);
    $("#v_mes_polizas").val(_mes);

    $("#v_moneda_poliza").val("");
    $("#v_moneda_poliza").change();

    $("#customFile").val("");

    $('#modal-sm-polizas').modal('show');
}

function mostrar_form_cobranzas() {
    $('#modal-sm-cobranzas').modal('show');
}

function mostrar_form_endosos() {
    $('#modal-sm-endosos').modal('show');
}


function GuardarUniversal() {
    var id = $("#v_id").val();
    var _periodo = $("#v_periodo_universal").val();
    
    if ($.trim(_periodo) == '') {
        toastr.error("Ingrese cantidad");
        return false;
    }

    if (id == "0") {
        grabaitem('Guardar');
    } else {
        grabaitem('Actualizar');
    }

}


function grabaitem(_accion) {

    var formData = $('#frmUniversalRegistrar').serialize();
    //console.log(formData);

    $.ajax({
        url: "/Universal/" + _accion,
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-universal").modal('hide');
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


function GuardarDataPoliza() {
    var _nombreid = "#customFile";
    var _moneda = $("#v_moneda_poliza").val();
    var fileUpload = $(_nombreid).get(0);
    var files = fileUpload.files;


    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    // Creando objeto FormData
    var fileData = new FormData();

    // Looping over all files and add it to FormData object  
    for (var i = 0; i < files.length; i++) {
        fileData.append(files[i].name, files[i]);
        //alert(files[i].name);
    }

    // Adding one more key to FormData object  
    fileData.append('moneda', _moneda);
    fileData.append('anio', $("#v_anio_polizas").val());
    fileData.append('mes', $("#v_mes_polizas").val());

    $.ajax({
        url: '/Universal/ExcelToDatabase',
        type: "POST",
        contentType: false, // Not to set any content header  
        processData: false, // Not to process data  
        data: fileData,
        success: function (data) {
            console.log(data);
            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-polizas").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }
        },
        error: function (err) {
            alert_error(err.statusText);
        }
    });

}


async function verDatosItem(_origen, _numreq, _item) {
    //limpiarFormulario();
    //$("#titulo").val("Actualizar Parametro");
    //$("#titulo_modal").html("Actualizar Comision");
    //alert(_id);

    var _id = $('#busca_intvisita2').val();
    const data = await myLeerVisita(_id);
    var _sw = 1;
    $.each(JSON.parse(data), function (key, item) {
        //alert(item.nombre_solicitante);
        if (item.ESTADOREQ == '2') {
            _sw = 0;
            alert_error("Imposible modificar", "Requerimiento cambio de estado");
        }

    });

    if (_sw == 1) {
        $.ajax({
            type: "POST",
            url: "/Itemizado/getData",
            data: {
                origen: _origen,
                numreq: _numreq,
                item: _item
            },
            datatype: "html",
            contentType: 'application/x-www-form-urlencoded',
            success: function (data) {
                //alert("retorna data");
                //console.log("JIM");
                //console.log(data);
                $.each(JSON.parse(data), function (key, item) {

                    $("#v_id").val("1");

                    $("#v_item").val(_item);
                    $("#v_numreq").val(_numreq);
                    $("#v_origen").val(_origen);


                    $("#v_tipo").val(item.CHRTIPO);
                    $("#v_tipo").change();


                    $("#v_codigo").val(item.VCHCODIGOPRODUCTO);
                    $("#v_chrcodproducto").val(item.CHRCODIGOPRODUCTO);

                    $("#v_glosa").val(item.VCHGLOSA);
                    $("#v_glosa2").val(item.VCHGLOSA);
                    $("#v_glosa_2").val(item.VCHGLOSA);

                    $("#v_cantidad").val(item.NUMCANTIDAD);


                    $("#v_grupo").val(item.VCHGRUPO);

                    var _div_orden = document.getElementById('m_orden');
                    _div_orden.style.display = 'block';
                    $("#v_orden").val(item.INTORDEN);


                    $("#v_criticidad").val(item.CHRCRITICIDAD);
                    $("#v_criticidad").change();


                    loadCombo(item.CHRUNIDAD, 'Itemizado', 'ListUnidades', 'v_chrunidad', 1);
                    //loadCombo(item.CHRTIPO, 'Avance', 'ListTipos', 'v_tipo', 1);
                    //loadCombo(item.id_presidente, 'Entidad', 'Lista_general', 'v_id_presidente', 1);

                    var _uni = document.getElementById('v_chrunidad');
                    var _glosa = document.getElementById('v_glosa');
                    if (item.CHRCODIGOPRODUCTO != "") {
                        _uni.disabled = true;
                        _glosa.disabled = true;
                    } else {
                        _uni.disabled = false;
                        _glosa.disabled = false;
                    }

                });

            },
            error: function () {
                toastr.error("Ocurrio un error!!");
            }
        });


        $('#modal-sm-itemizado').modal('show');
    }

}


async function InactivarItemi(_origen, _numreq, _item) {

    var _id = $('#busca_intvisita2').val();
    const data = await myLeerVisita(_id);
    var _sw = 1;
    $.each(JSON.parse(data), function (key, item) {
        //alert(item.nombre_solicitante);
        if (item.ESTADOREQ == '2') {
            _sw = 0;
            alert_error("Imposible inactivar", "Requerimiento cambio de estado");
        }

    });

    if (_sw == 1) {
        swal({
            title: "¿Seguro(a) de inactivar itemizado?",
            text: "REQ. N° " + _numreq,
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
                    cambiarestado(_origen, _numreq, _item, 'Activo');
                } else {
                    swal("Operación cancelada");
                }
            });
    }

}

function ActivarItemi(_origen, _numreq, _item) {
    swal({
        title: "¿Seguro(a) de activar itemizado?",
        text: "REQ. N° " + _numreq,
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
                cambiarestado(_origen, _numreq, _item, 'Inactivo');
            } else {
                swal("Operación cancelada");
            }
        });
}


function cambiarestado(_origen, _numreq, _item, _accion) {
    //alert("inactivo");
    //var _id = $("#v_idina").val();
    //alert("activar " + _id);

    $.ajax({
        url: "/Itemizado/" + _accion,
        type: "POST",
        data: {
            origen: _origen,
            numreq: _numreq,
            item: _item
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