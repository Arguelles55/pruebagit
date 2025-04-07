<?php

use kartik\select2\Select2;
use yii\web\JsExpression;
use yii\helpers\Html;
use yii\web\View;

$url_fechas_historico = Yii::$app->urlManager->createUrl(['consulta/get-fecha-certificados']);
$url_consulta_historico = Yii::$app->urlManager->createUrl(['consulta/historico-proveedor']);
$url_descarga_historico = Yii::$app->urlManager->createUrl(['consulta/descargar-reporte-historico']);
$url_descarga_historial = Yii::$app->urlManager->createUrl(['consulta/descargar-historial-certificacion']);

$url_proveedor_consulta = Yii::$app->urlManager->createUrl(['consulta/get-proveedor-historico']);



$this->registerJs("

let provider_select_id = '#proveedor_select';
let fecha_select_id = '#fecha_select';
let container_data_id = '#ajax-contenido';

let btn_historico = '.descarga-historico';
let btn_historial = '.descarga-historial';

$(document).ready(() => {

    function swAlrtError(message = 'Ocurrio un error', titlo = 'Algo salio mal!'){
        swal('Consulta incompleta', message, 'warning' );
    }

    //Evento en Select2 (Proveedor)
    $(provider_select_id).change(function(event){
        let proveedor_id = $(this).val();
        if( proveedor_id ){
            $.get('{$url_fechas_historico}?provider_id='+proveedor_id, function(data) { $(fecha_select_id).html(data); } );
        }
    });

    $(fecha_select_id).change(function(event){
        let timestamp = $(this).val();
        let proveedor_id = $(provider_select_id).val();
        if( timestamp && proveedor_id ){
            fireHoldOn();
            $.get('{$url_consulta_historico}',{ provider_id: proveedor_id, time: timestamp })
            .done( ( data ) => { $(container_data_id).html(data); })
            .always( () => { HoldOn.close(); });
        }
    });

    $(btn_historico).click( (event) => {
        let timestamp = $(fecha_select_id).val();
        let proveedor_id = $(provider_select_id).val();

        if( (timestamp === undefined || timestamp == '') || 
            (proveedor_id === undefined || proveedor_id == '') ){ swAlrtError(`Debe seleccionar un proveedor y una fecha de certificacion, de faltar alguno no podra realizar la consulta`) }
        else{

            let title = 'Confirmar';
            let text = `El archivo XLSX será generado para descargar con todos los datos de la consulta, esto puede tomar tiempo.
                        Deshabilite cualquier bloqueador de ventanas emergentes, para una descarga adecuada`;
            let icon = 'warning';
            let confirm_label = 'De acuerdo';
            let cancel_label = 'Cancelar';
            let url = '{$url_descarga_historico}?provider_id='+proveedor_id+'&time='+timestamp;

            swal({
                title: title,
                text: text,
                icon: icon,
                buttons: [cancel_label, confirm_label],
            }).then( isConfirm => {
                if( isConfirm ){ fireHoldOn(4000); window.location.href = url;}
            });

        }

    });

    $(btn_historial).click( (event) => {
        let proveedor_id = $(provider_select_id).val();

        if( proveedor_id === undefined || proveedor_id == '' ){ swAlrtError(`Debe seleccionar un proveedor para poder descargar su historial de certificaciones`) }
        else{

            let title = 'Confirmar';
            let text = `El archivo ZIP será generado para descargar con todos los datos del historico del proveedor seleccionado, esto puede tomar tiempo.
                        Deshabilite cualquier bloqueador de ventanas emergentes, para una descarga adecuada`;
            let icon = 'warning';
            let confirm_label = 'De acuerdo';
            let cancel_label = 'Cancelar';
            let url = '{$url_descarga_historial}?provider_id='+proveedor_id;

            swal({
                title: title,
                text: text,
                icon: icon,
                buttons: [cancel_label, confirm_label],
            }).then( isConfirm => {
                if( isConfirm ){ fireHoldOn(6000); window.location.href = url;}
            });

        }

    });
})



");

$this->registerJs("
    const mapResponse = (results) => {
        let mapped = results.map( (obj) => {
            return { 'id' : obj.provider_id, 'text' : obj.fullname };
        });
        return { 'results' : mapped };
    }

    const queryProveedor = (params) => {
        return { texto : params.term };
    };


", View::POS_HEAD);

?>
<div class="contenedor_general">
    <div class="caja_titulo" style="display: flex;">
        <span>Consulta historico de proveedores</span>
        <div class="head-flex-btns">
            <?= Html::a('<span class="glyphicon glyphicon-download"></span>  Descargar (historico)', false,
                [ 'class' => 'btn btn-secondary descarga-historico' ] ); ?>
            <?= Html::a('<span class="glyphicon glyphicon-download"></span>  Descargar historial (historico)', false,
                [ 'class' => 'btn btn-secondary descarga-historial' ] ); ?>
        </div>
    
    </div>
    <div class="row">
        <div class="col-md-6 col-12">
            <?= Select2::widget([
                'id' => 'proveedor_select',
                'name' => 'proveedor',
                'data' => [],
                'options' => [ 'placeholder' => 'Busqueda por Nombre/Razon Social/RFC', 'multiple' => false ],
                'pluginOptions' => [
                    'minimumInputLength' => 1,
                    'language' => [
                        'errorLoading' => new JsExpression("() => { return 'Cargando...'; }"),
                    ],
                    'ajax' => [
                        'url' => $url_proveedor_consulta,
                        'dataType' => 'json',
                        'delay' => 250,
                        'data' => new JsExpression("(params) => { return queryProveedor(params); } "),
                        'processResults' => new JsExpression("(data) => { return mapResponse(data); }")
                    ],
                ]    
            ])?>
        </div>
        <div class="col-md-6 col-12">
            <?= Select2::widget([
                'name' => 'fecha',
                'data' => [],
                'options' => [
                    'id' => 'fecha_select',
                    'placeholder' => 'Selecciona fecha',
                    'multiple' => false
                ]
            ]) ?>
        </div>
    </div>

    <div id='ajax-contenido'></div>

</div>
