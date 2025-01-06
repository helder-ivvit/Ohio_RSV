
var __filas = 15;

loadDataConfig();

function loadDataConfig() {
    var totregistros = 0;
    var paginas = 0;
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
