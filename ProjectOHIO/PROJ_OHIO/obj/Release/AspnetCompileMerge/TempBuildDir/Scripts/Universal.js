
var __filas = 15;

loadDataUniversal();

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

    var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick="IrResultadoVP(\'' + _anio + '\',\'' + _mes + '\')"><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
    var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick="IrResultadoVU(\'' + _anio + '\',\'' + _mes + '\')"><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matem&aacute;tica VU</a>';

    if (_activo == 1) {
        _anio_html = '<font color="#000000">' + _anio + '</font>';
        _mes_html = '<font color="#000000">' + _nombre_mes + '</font>';

        var _opt_polizas = '<a href="#" class="dropdown-item" onclick="mostrar_form_polizas(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de P&oacute;lizas</a>';
        var _opt_cobranzas = '<a href="#" class="dropdown-item" onclick="mostrar_form_cobranzas(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Cobranzas</a>';
        var _opt_endosos = '<a href="#" class="dropdown-item" onclick="mostrar_form_endosos(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-database"></i>&nbsp;&nbsp;&nbsp;Cargar Data de Endosos</a>';
        var _opt_valor_poliza = '<a href="#" class="dropdown-item" onclick="CalcularValorPoliza(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Valor P&oacute;liza</a>';
        var _opt_vida_univ = '<a href="#" class="dropdown-item" onclick="CalcularReserva(\'' + _anio + '\',\'' + _mes + '\',\'' + _nombre_mes + '\')"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Reserva Matem&aacute;tica VU</a>';
        var _opt_vida_univ_moce = '<a href="#" class="dropdown-item" onclick="CalcularVUMoce()"><i class="fas fa-cogs"></i>&nbsp;&nbsp;&nbsp;Procesar Reserva Matem&aacute;tica MOCE</a>';
        
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
        _opt = _opt + '        ' + _opt_vida_univ_moce;
        _opt = _opt + '    </div>';
        _opt = _opt + '</div>';

    }

    if (_activo == 0) {
        // #c0392b
        _anio_html = '<font color="#ffa000">' + _anio + '</font>';
        _mes_html = '<font color="#ffa000">' + _nombre_mes + '</font>';

        //var _opt_result_valor_poliza = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Valor P&oacute;liza</a>';
        //var _opt_result_vida_univ = '<a href="#" class="dropdown-item" onclick=""><i class="fas fa-clone"></i>&nbsp;&nbsp;&nbsp;Resultados Reserva Matematica VU</a>';

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

function CalcularVUMoce(_anio, _mes, _nombremes) {
    $('#modal-sm-procesar_vu_moce').modal('show');
}

function CalcularValorPoliza(_anio, _mes, _nombremes) {

    $("#titulo_modal_valor_poliza").html('Procesar Valor Poliza <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(' + _nombremes + '-' + _anio + ')</b>');

    $("#v_anio_vp").val(_anio);
    $("#v_mes_vp").val(_mes);

    $("#v_moneda_vp").val("");
    $("#v_moneda_vp").change();

    $('#modal-sm-procesar_vp').modal('show');

    //swal({
    //    title: "¿Seguro de procesar el Valor Poliza del periodo " + nombremes + "-" + anio + "?",
    //    text: "",
    //    icon: "warning",
    //    buttons: true,
    //    dangerMode: true,
    //    buttons: ['CANCELAR', 'ACEPTAR']
    //})
    //    .then((willDelete) => {
    //        if (willDelete) {
                
    //            //cambiarestado(_anio, _mes, 'Activo');
    //            alert("Procesar valor poliza");

    //        } else {
    //            swal("Operación cancelada");
    //        }
    //    });

}


function CalcularReserva(_anio, _mes, _nombremes) {

    $("#titulo_modal_vu").html('Procesar Reserva Matemática V.U. <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(' + _nombremes + '-' + _anio + ')</b>');

    $("#v_anio_reserva_vu").val(_anio);
    $("#v_mes_reserva_vu").val(_mes);

    $("#v_moneda_reserva_vu").val("");
    $("#v_moneda_reserva_vu").change();

    $('#modal-sm-procesar_vu').modal('show');

}


async function limpiarFormularioUniversal() {

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


function mostrar_form_cobranzas(_anio, _mes, _mesnombre) {

    $("#titulo_modal_cobranza").html('Cargar Data de Cobranzas <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(Periodo ' + _anio + '-' + _mesnombre + ')</b>');

    $("#v_anio_cobranza").val(_anio);
    $("#v_mes_cobranza").val(_mes);

    $("#v_moneda_cobranza").val("");
    $("#v_moneda_cobranza").change();

    $('#modal-sm-cobranzas').modal('show');
}


function mostrar_form_endosos(_anio, _mes, _mesnombre) {

    $("#titulo_modal_endoso").html('Cargar Data de Endosos <b style="color: black;">&nbsp;&nbsp;&nbsp;&nbsp;(Periodo ' + _anio + '-' + _mesnombre + ')</b>');

    $("#v_anio_endoso").val(_anio);
    $("#v_mes_endoso").val(_mes);

    $('#modal-sm-endosos').modal('show');
}


function GuardarCobranza() {
    //var id = $("#v_id").val();
    var _moneda = $("#v_moneda_cobranza").val();

    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-cobranzas").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmCobranzaRegistrar').serialize();
    //console.log(formData);

    $.ajax({
        url: "/Universal/Cobranza",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-cobranzas").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });

}


function GuardarEndoso() {
    //var id = $("#v_id").val();

    $("#modal-sm-endosos").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmEndosoRegistrar').serialize();
    //console.log(formData);

    $.ajax({
        url: "/Universal/Endoso",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-endosos").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);                
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });

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

    $("#modal-sm-polizas").modal('hide');
    $('#modal-loading').modal('show');

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
                $('#modal-loading').modal('hide');
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

function GuardarVP() {    
    var _moneda = $("#v_moneda_vp").val();
    
    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-procesar_vp").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmProcesaValorPoliza').serialize();
    //console.log(formData);
    //alert("123");
    $.ajax({
        url: "/Universal/ProcesaVP",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-procesar_vp").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });
}


function IrResultadoVP(_anio, _mes) {
    const url = $("#RedirectResultadoVP").val();
    const new_url = url.replace("100_", _anio); /// modificar valores
    const new_url2 = new_url.replace("200_", _mes);
    //console.log("final  " + new_url2);
    location.href = new_url2;
}


function GuardarReservaVU() {
    var _moneda = $("#v_moneda_reserva_vu").val();

    if ($.trim(_moneda) == '') {
        toastr.error("Seleccione moneda");
        return false;
    }

    $("#modal-sm-procesar_vu").modal('hide');
    $('#modal-loading').modal('show');

    var formData = $('#frmProcesaReservaMatematicaVU').serialize();
    //console.log(formData);
    //alert("123");
    $.ajax({
        url: "/Universal/ProcesaVU",
        type: "POST",
        data: formData,
        contentType: 'application/x-www-form-urlencoded; charset=UTF-8', // optional
        success: function (data) {

            //console.log(data);

            if (data.operacion == true) {
                toastr.success(data.mensaje);

                loadDataUniversal();
                $("#modal-sm-procesar_vu").modal('hide');
            }
            else {
                toastr.error(data.mensaje);
                //alert("Ocurrio un error!!");
            }

            if (data.operacion2 == false) {
                toastr.error(data.mensaje2);
            }

            $('#modal-loading').modal('hide');

        },
        error: function () {
            toastr.error("Ocurrio un error!!");
            $('#modal-loading').modal('hide');
            //alert("Ocurrio un error!!");
        }
    });
}


function IrResultadoVU(_anio, _mes) {
    const url = $("#RedirectResultadoVU").val();
    const new_url = url.replace("100_", _anio); /// modificar valores
    const new_url2 = new_url.replace("200_", _mes);
    //console.log("final  " + new_url2);
    location.href = new_url2;
}




