<?php
namespace app\controllers;

use app\helpers\GeneralController;
use app\helpers\SireController;
use app\models\DatosValidados;
use app\models\FotografiaNegocio;
use app\models\Giro;
use app\models\Historico;
use app\models\HistoricoCartaProtesta;
use app\models\HistoricoCertificados;
use app\models\HistoricoContrataciones;
use app\models\Provider;
use app\models\ProviderGiro;
use app\models\ProviderQuery;
use app\models\Sanciones;
use app\models\Ubicacion;
use DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Yii;
use yii\base\Controller;
use yii\db\Expression;
use yii\helpers\ArrayHelper;

class ConsultaController extends GeneralController
{
    public function actionSire($search=null, $type='rfc')
    {
       if(!$search)
           return $this->render('sire');


       $data = ["results"=>[
           SireController::consultaProveedor($search,$type)
       ]];

       //busqueda

        return $this->render('sire',$data);

    }

    public function actionIndexFuncionarios(){

        $params = Yii::$app->request->queryParams;
        $dataProvider = ProviderQuery::getFuncionariosData($params);

        return $this->render('index-funcionarios', [
            'dataProvider' => $dataProvider,
        ]);

    }

    //Actualiza la DB con los datos de los funcionarios, ejecutar en CRON ya que tarda apox 25 min
    public function actionExportAllFuncionarios(){

        $unique_rfc = ProviderQuery::updateFuncionariosData();

        $post_data = json_encode( array_values ($unique_rfc) );
        $response_consulta = GeneralController::restConsultaFuncionarios($post_data);

        if( !is_null($response_consulta) ) {
            foreach($response_consulta as $respuesta){
                $es_funcionario = $respuesta['respuesta'] == "Si es funcionario" ? "true"  : "false";
                Yii::$app->db->createCommand("UPDATE public.funcionarios_proveedores SET es_funcionario = $es_funcionario WHERE funcionario_rfc = '".$respuesta['rfc']."'")->execute();
            }
        }

        echo "Ok";

        return true;
    }

    public function actionHistoricoProveedor($provider_id, $time){
        $provider = Provider::findOne($provider_id);

        $date = new DateTime();
        $date->setTimestamp($time);
        $date_string = date_format($date, 'Y-m-d');

        $permisos = Yii::$app->db->createCommand("SELECT distinct(module) FROM permission_view_module")->queryAll();
        $modulos_permisos = array_unique(ArrayHelper::getColumn($permisos,'module'));

        $searchModelCarta = new \app\models\HistoricoCartaProtestaSearch();
        $searchModelCarta->provider_id = $provider->provider_id;
        $dataProviderCarta = $searchModelCarta->search(Yii::$app->request->queryParams,'consultaCartas');

        $searchModelCertificado = new \app\models\HistoricoCertificadosSearch();
        $searchModelCertificado->provider_id = $provider->provider_id;

        $dataProviderCertificado = $searchModelCertificado->search(Yii::$app->request->queryParams, 'validCertificados');
        $dataProviderCursos = $searchModelCertificado->search(Yii::$app->request->queryParams, 'validCursos');

        $filtro_certificados = HistoricoCertificados::find()->select( new Expression('created_at::date as date') )->where( ['and', ['provider_id' => $provider->provider_id], ['tipo' => 'CERTIFICADO'], ['provider_type' => 'bys'], ['<', 'date(created_at)', $date_string] ] )->orderBy(['created_at' => SORT_DESC])->asArray()->all();
        $fechas_periodos = ArrayHelper::getColumn($filtro_certificados,'date');

        $array_response = [
            'cartas_data' => $dataProviderCarta,
            'historico_cursos' => $dataProviderCursos,
            'historico_certificados_data' => $dataProviderCertificado,
            'modulos_permisos' => $modulos_permisos,
            'type' => ''
        ];

        $array_response = array_merge($array_response, self::getHistoricoData($provider, $date_string, $fechas_periodos));
        return $this->renderPartial('../provider/detalle-consulta-permisos', $array_response );

    }

    public function actionGetProveedoresHistorico(){
        $proveedores_id_data = Yii::$app->db->createCommand("
            SELECT p.provider_id, CASE WHEN (p.name_razon_social is not null and p.name_razon_social != '')
            THEN p.name_razon_social ELSE concat_ws(' ', p.pf_nombre, p.pf_ap_paterno, p.pf_ap_materno) END as fullname
            FROM provider p WHERE provider_id in ( SELECT distinct(h.provider_id) FROM provider.historico h WHERE h.tipo='bys' ) ORDER BY fullname")->queryAll();
        
        $proveedores = ArrayHelper::map($proveedores_id_data,'provider_id', 'fullname');

        return json_encode($proveedores);
    }

    /**
     * Funcion que retorna las fechas de certificacion del proveedor
     * @param int $provider_id Id del proveedor a consultar
     * @return string HTML con las fechas de certificacion encontradas dentro de etiquetas <option>
     */
    public function actionGetFechaCertificados($provider_id){
        $resultado = Yii::$app->db->createCommand("SELECT date FROM ( SELECT DISTINCT(created_at::date) as date FROM historico_certificados 
            WHERE provider_id = $provider_id AND tipo = 'CERTIFICADO' AND provider_type = 'bys' ) x ORDER BY date DESC ")->queryAll();
        $fechas = ArrayHelper::getColumn($resultado,'date');
        $options = '<option value="">Selecciona...</option>';
        foreach($fechas as $fecha){
            $dateObj = new DateTime($fecha);
            $timeKey = $dateObj->getTimestamp();
            $options .= "<option value=".$timeKey.">$fecha</option>";
        }
        return $options;
    }

    public function actionConsultaHistoricoProveedores(){
        $proveedores_id_data = Yii::$app->db->createCommand("
        SELECT p.provider_id, CASE WHEN (p.name_razon_social is not null and p.name_razon_social != '')
        THEN p.name_razon_social ELSE concat_ws(' ', p.pf_nombre, p.pf_ap_paterno, p.pf_ap_materno) END as fullname
        FROM provider p WHERE provider_id in ( SELECT distinct(h.provider_id) FROM provider.historico h WHERE h.tipo='bys' ) ORDER BY fullname")->queryAll();
    
        $proveedores = ArrayHelper::map($proveedores_id_data,'provider_id', 'fullname');
        return $this->render('view_consulta',['proveedores' => $proveedores]);
    }

    public function actionGetProveedorHistorico($texto){
        $query = "SELECT p.provider_id, CASE WHEN  p.tipo_persona = 'Persona moral' THEN p.name_razon_social ELSE concat_ws(' ', p.pf_nombre, p.pf_ap_paterno, p.pf_ap_materno) END as fullname
        FROM provider p LEFT JOIN provider.rfc r ON r.provider_id = p.provider_id WHERE ( 
                (p.tipo_persona = 'Persona moral' AND unaccent(p.name_razon_social) ILIKE unaccent('%{$texto}%')) 
                OR 
                (p.tipo_persona != 'Persona moral' AND unaccent(concat_ws(' ', p.pf_nombre, p.pf_ap_paterno, p.pf_ap_materno)) ILIKE unaccent('%{$texto}%'))
                OR
                r.rfc ILIKE '%{$texto}%'
              )
              AND p.provider_id IN (SELECT DISTINCT(h.provider_id) FROM provider.historico h WHERE h.tipo = 'bys') ORDER BY fullname LIMIT 10";

        $return_query = Yii::$app->db->createCommand($query)->queryAll();
        return json_encode($return_query);
    }

    /**
    * Funcion de encapsulamiento de datos del historico para reutilizarse en la consulta y llenado del reporte
    */
    public function getHistoricoData($provider_data, $fecha_filtro, $fechas_certificados){
        return [
            'nombre_provider' => $provider_data->getFullName(),
            'provider_data' => $provider_data,
            'comprobante_domicilio' => self::getHistoricoModelo('ComprobanteDomicilio', $provider_data->provider_id, $fecha_filtro),
            'rfc_data' => self::getHistoricoModelo('Rfc', $provider_data->provider_id, $fecha_filtro),
            'acta_data' => self::getHistoricoModelo('ActaConstitutiva', $provider_data->provider_id, $fecha_filtro),
            'mod_acta_data' => self::getArrayHistorico('ModificacionActa', $provider_data->provider_id, $fecha_filtro, $fechas_certificados), //self::getHistoricoModelo('ModificacionActa', $provider_id, true),
            'model_acta_constitutiva' => self::getHistoricoModelo('ActaConstitutiva', $provider_data->provider_id, $fecha_filtro),
            'representante_data' => self::getArrayHistorico('RepresentanteLegal', $provider_data->provider_id, $fecha_filtro, $fechas_certificados), //self::getHistoricoModelo('RepresentanteLegal', $provider_id, true),
            'accionistas_data' => self::getArrayHistorico('RelacionAccionistas', $provider_data->provider_id, $fecha_filtro, $fechas_certificados), //self::getHistoricoModelo('RelacionAccionistas', $provider_id, true),
            'actividades_data' => self::getHistoricoGiro($provider_data->provider_id, $fecha_filtro),
            'productos_data' => ProviderController::getHistoricoProviderGiro($provider_data->provider_id),
            'ult_declaracion_data' => self::getHistoricoModelo('UltimaDeclaracion', $provider_data->provider_id, $fecha_filtro),
            'clientes_data' => self::getArrayHistorico('ClientesContratos', $provider_data->provider_id, $fecha_filtro, $fechas_certificados), //self::getHistoricoModelo('ClientesContratos', $provider_id, true),
            'certificados_data' => self::getArrayHistorico('Certificacion', $provider_data->provider_id, $fecha_filtro, $fechas_certificados), //self::getHistoricoModelo('Certificacion', $provider_id, true),
            'ubicacion_fiscal_data' => self::getUbicacionUnicaHistorico($provider_data->provider_id, 'DOMICILIO FISCAL', $fecha_filtro, $fechas_certificados, true),
            'ubicacion_nl_data' => self::getUbicacionUnicaHistorico($provider_data->provider_id, 'DIRECCIÓN NUEVO LEÓN', $fecha_filtro, $fechas_certificados),
            'ubicaciones_data' => self::getArrayUbicacionesHistorico($provider_data->provider_id, $fecha_filtro, $fechas_certificados),
            'ubicaciones_arr' => []/* self::getArrayHistoricoCordUbicaciones($provider_id) */,
            'pago_data' => self::getHistoricoModelo('IntervencionBancaria', $provider_data->provider_id, $fecha_filtro),
            'curp_data' => self::getHistoricoModelo('Curp', $provider_data->provider_id, $fecha_filtro),
            'idOf_data' => self::getHistoricoModelo('IdOficial', $provider_data->provider_id, $fecha_filtro),
            'perfil' => self::getHistoricoModelo('Perfil', $provider_data->provider_id, $fecha_filtro),
            'provider_ext' => self::getHistoricoModelo('Provider', $provider_data->provider_id, $fecha_filtro),
        ];
    }

    /**
     * Funcion que mapea los datos del historico a un objeto de Yii (Version de consulta por certificacion)
     * @param string $modelo Nombre del modelo registrado en Yii en el namespace models
     * @param int $provider_id Id del proveedor a realizar la consulta
     * @param string $fecha Fecha de certificacion en formato yyyy-mm-dd, retornara el primer registro igual o menor a la fecha establecida
     * @return Instance Instancia del modelo, sera vacia si no hay registro pero no null
     */
    public function getHistoricoModelo($modelo, $provider_id, $fecha, $tipo = 'bys'){
        $class = "app\models\\" . $modelo; //Prepara la clase del modelo a instanciar
        $opciones_consulta = ['and',['modelo' => $modelo], ['provider_id' => $provider_id], ['tipo' => $tipo] ]; //Opciones de consulta default
        if(!is_null($fecha) && !empty($fecha)){ array_push($opciones_consulta, ['<=','date(fecha_validacion)', $fecha]); } //Añade la condicion si existe una fecha filtro
        $model = Historico::find()->where($opciones_consulta)->orderBy(['fecha_validacion' => SORT_DESC])->one();
        if(is_null($model) || empty($model)){ return new $class(); }
        if(GeneralController::is_multidimentional($model->data) && $modelo != 'Ubicacion'){
            $response = self::setAtributesModel($class, $model->data[0]); 
        }else{ $response = self::setAtributesModel($class, $model->data); }

        return $response;
    }

    public function getArrayHistorico($modelo, $provider_id, $fecha, $fechas_certificados,$tipo='bys'){
        $response = null;
        if( !empty($fechas_certificados) ){
            foreach($fechas_certificados as $fecha_filtro){
                $response = self::getPeriodoArrayHistorico($modelo, $provider_id, $fecha, $fecha_filtro, $tipo, 'fecha_validacion');
                if(!is_null($response)){ continue; }
            }
        }
        return is_null($response) ? self::getUltimoArrayHistorico($modelo, $provider_id, $fecha, $tipo) : $response;
    }

    public function getArrayUbicacionesHistorico($provider_id, $fecha, $fechas_certificados){
        $response = null;
        if( !empty($fechas_certificados) ){
            foreach($fechas_certificados as $fecha_filtro){
                $response = self::getUbicacionesHistorico($provider_id, $fecha, $fecha_filtro);
                if(!is_null($response) && !empty($response)){ continue; }
            }
        }
        return is_null($response) || empty($response) ? self::getUbicacionesHistorico($provider_id, $fecha) : $response;
    }

    public function getUltimoArrayHistorico($modelo, $provider_id, $fecha, $tipo){
        $class = "app\models\\" . $modelo; //Prepara la clase del modelo a instanciar
        $response = [new $class];
        $opciones_consulta = ['and',['modelo' => $modelo], ['provider_id' => $provider_id], ['tipo' => $tipo],["json_typeof(data)" => 'array'] ]; //Opciones de consulta default
        if(!is_null($fecha) && !empty($fecha)){ array_push($opciones_consulta, ['<=','date(fecha_validacion)', $fecha]); } //Añade la condicion si existe una fecha filtro
        $model_v3 = Historico::find()->where($opciones_consulta)->one(); //De ser el registro mas reciente de la ultima version, regresara un arreglo con todos los datos validados
        if( is_null($model_v3) ){
            $opciones_consulta[4]['json_typeof(data)'] = 'object'; //Si no hay array en consulta, se cambia la consulta a objetos de V1
            $model_v1 = Historico::find()->where($opciones_consulta)->all(); //Posible agregar otro filtro
            if( count($model_v1) > 0 ){
                $response = [];
                foreach( $model_v1 as $historico ){
                    $instancia = self::setAtributesModel($class, $historico->data);
                    array_push($response, $instancia);
                }
            }
        }else{
            $response = [];
            foreach( (array) $model_v3->data as $historico ){
                $result = self::setAtributesModel($class, $historico);
                array_push($response, $result);
            }
        }

        return $response;
    }

    public function getPeriodoArrayHistorico($modelo, $provider_id, $fecha_inicio, $fecha_fin, $tipo, $order_field){
        $class = "app\models\\" . $modelo; //Prepara la clase del modelo a instanciar
        $response = null;
        $opciones_consulta = ['and',['modelo' => $modelo], ['provider_id' => $provider_id], ['tipo' => $tipo],
            ["json_typeof(data)" => 'array'], ['<=','date(fecha_validacion)', $fecha_inicio], ['>','date(fecha_validacion)', $fecha_fin] ]; //Opciones de consulta default
        $model_v3 = Historico::find()->where($opciones_consulta)->orderBy([
            $order_field => SORT_DESC
          ])->one(); //De ser el registro mas reciente de la ultima version, regresara un arreglo con todos los datos validados
        if( is_null($model_v3) ){
            $opciones_consulta[4]['json_typeof(data)'] = 'object'; //Si no hay array en consulta, se cambia la consulta a objetos de V1
            $model_v1 = Historico::find()->where($opciones_consulta)->all(); //Posible agregar otro filtro
            if( count($model_v1) == 0 ){ return $response; }
            $response = [];
            foreach( $model_v1 as $historico ){
                $instancia = self::setAtributesModel($class, $historico->data);
                array_push($response, $instancia);
            }
        }else{
            $response = [];
            foreach( (array) $model_v3->data as $historico ){
                $result = self::setAtributesModel($class, $historico);
                array_push($response, $result);
            }
        }

        return $response;
    }

    public function getHistoricoGiro($provider_id, $fecha, $tipo = 'bys'){
        $response = [];
        self::setTempHistoricoGiro($provider_id, $fecha, $tipo);

        $model = Giro::find()->from('temp_giro_historico')->where(['provider_id' => $provider_id])->all();
        $response = ($model == null || empty($model)) ? [new Giro()] : $model;

        return $response;
    }

    public function setTempHistoricoGiro($provider_id, $fecha, $tipo){
        Yii::$app->db->createCommand("DROP TABLE IF EXISTS temp_giro_historico")->execute();

        Yii::$app->db->createCommand(
            "CREATE TEMPORARY TABLE temp_giro_historico(
                giro_id bigint, actividad_id bigint,
                provider_id bigint, porcentaje smallint,
                start_date date,
                url_factura_c_c text
            )" )->execute();

        //EL JSON de la columna data en Historicos la convierte en un solo array de objetos con el fin de pasarlos despues a registros en la DB
        //Agrega un nuevo campo al json object que indica si el producto es legacy de acuerdo a la fecha de validacion
        $query_json_line = "SELECT jsonb_agg( arr.item_object ) FROM provider.historico, json_array_elements(data)
                            with ordinality arr(item_object) WHERE modelo = 'Giro' AND provider_id = $provider_id AND tipo = '$tipo'  AND fecha_validacion::date <= '$fecha'
                            GROUP BY fecha_validacion ORDER BY fecha_validacion DESC LIMIT 1";

        //EL JSON array se convierte en registros para poder ser manipulados con SQL
        $query_json_record = "SELECT * FROM jsonb_to_recordset( ($query_json_line) ) as x(
                    giro_id bigint, actividad_id bigint, provider_id bigint, porcentaje smallint, start_date date, url_factura_c_c text ) ";

        //Se insertan los registros en una tabla temporal para manipularlos con querys
        $query_json_insert_temp = "INSERT INTO temp_giro_historico ($query_json_record)";

        Yii::$app->db->createCommand($query_json_insert_temp)->execute();
    }

    public function getUbicacionUnicaHistorico($provider_id, $tipoUbicacion, $fecha, $fechas_certificados, $isRequired = false, $tipo = 'bys'){
        $response = ["ubicacion" => $isRequired ? new Ubicacion() : null, "fotografia" => $isRequired ? [ new FotografiaNegocio() ] : null ];
        //Busqueda de modelo V3 de la ubicacion unica
        $ubicacion_json = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                    WHERE modelo = 'Ubicacion' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'tipo' = :tipo_ubicacion 
                    AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha ORDER BY fecha_validacion DESC",
                    ['provider_id' => $provider_id, 'tipo' => $tipo, 'tipo_ubicacion' => $tipoUbicacion, 'fecha' => $fecha])->queryOne();
        if( empty($ubicacion_json['item_object']) ){
            $model = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'Ubicacion' AND provider_id = $provider_id 
                    AND tipo = '$tipo' AND data::json->>'tipo' = '$tipoUbicacion' AND json_typeof(data) = 'object' AND fecha_validacion::date <= '$fecha'")
                    ->orderBy(['fecha_validacion' => SORT_DESC])->one();
            if( !empty($model) ){
                $response["ubicacion"] = self::setAtributesModel("app\models\\Ubicacion", $model->data);
                $response["fotografia"] = self::getModeloArrayFotografia($provider_id, $response["ubicacion"]->ubicacion_id, $fecha, $fechas_certificados, 'V1');
            }
        }else{
            $response["ubicacion"] = self::setAtributesModel("app\models\\Ubicacion", json_decode($ubicacion_json['item_object']));
            $response["fotografia"] = self::getModeloArrayFotografia($provider_id, $response["ubicacion"]->ubicacion_id, $fecha, $fechas_certificados, 'V3');
        }

        return $response;
    }

    public function getModeloArrayFotografia($provider_id, $ubicacion_id, $fecha_consulta, $fechas_certificados, $modo = 'V1'){
        $response = null;
        if( !empty($fechas_certificados) ){
            foreach($fechas_certificados as $fecha_filtro){
                $response = $modo == 'V1' ? self::getFotografiasUbicacionV1($provider_id, $ubicacion_id, $fecha_consulta, $fecha_filtro)
                    : self::getFotografiasUbicacionV3($provider_id, $ubicacion_id, $fecha_consulta, $fecha_filtro);
                if(!is_null($response)){ continue; }
            }
        }
        return is_null($response) ? ( $modo == 'V1' ? self::getFotografiasUbicacionV1($provider_id, $ubicacion_id, $fecha_consulta)
            :  self::getFotografiasUbicacionV3($provider_id, $ubicacion_id, $fecha_consulta) ) : $response;
    }

    public function getFotografiasUbicacionV1($provider_id, $ubicacion_id, $fecha_consulta, $fecha_limite = null){
        $response = null;
        if( is_null($fecha_limite) ){
            $histFotografia = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'FotografiaNegocio' 
                    AND provider_id = $provider_id AND tipo = 'bys' AND data::json->>'ubicacion_id' = :ubicacion_id AND json_typeof(data) = 'object' 
                    AND fecha_validacion::date <= :fecha_consulta",
                    ['ubicacion_id' => $ubicacion_id, 'fecha_consulta' => $fecha_consulta])
                    ->orderBy(['fecha_validacion' => SORT_DESC])->all();
        }else{
            $histFotografia = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'FotografiaNegocio' 
                    AND provider_id = $provider_id AND tipo = 'bys' AND data::json->>'ubicacion_id' = :ubicacion_id AND json_typeof(data) = 'object' 
                    AND fecha_validacion::date <= :fecha_consulta AND fecha_validacion::date > :fecha_limite",
                    ['ubicacion_id' => $ubicacion_id, 'fecha_consulta' => $fecha_consulta, 'fecha_limite' => $fecha_limite ])
                    ->orderBy(['fecha_validacion' => SORT_DESC])->all();
        }
        if( !empty($histFotografia) ){ 
            $auxFoto = [];
            foreach($histFotografia as $fotoHist){
                array_push($auxFoto, self::setAtributesModel("app\models\\FotografiaNegocio", $fotoHist->data));
            }
            $response = $auxFoto; 
        }

        return $response;
    }

    public function getFotografiasUbicacionV3($provider_id, $ubicacion_id, $fecha_consulta, $fecha_limite = null){
        $response = null;
        if( is_null($fecha_limite) ){
            $histFotografia = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                WHERE modelo = 'FotografiaNegocio' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'ubicacion_id' = :ubicacion_id
                AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta ORDER BY fecha_validacion DESC",
                ['provider_id' => $provider_id, 'tipo' => 'bys', 'ubicacion_id' => $ubicacion_id, 'fecha_consulta' => $fecha_consulta])->queryOne();
        }else{
            $histFotografia = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                WHERE modelo = 'FotografiaNegocio' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'ubicacion_id' = :ubicacion_id
                AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta AND fecha_validacion::date > :fecha_limite ORDER BY fecha_validacion DESC",
                ['provider_id' => $provider_id, 'tipo' => 'bys', 'ubicacion_id' => $ubicacion_id, 'fecha_consulta' => $fecha_consulta, 'fecha_limite' => $fecha_limite ])->queryOne();
        }
        if( !empty($histFotografia['item_object']) ){ 
            $response = [ self::setAtributesModel("app\models\\FotografiaNegocio", json_decode($histFotografia['item_object'])) ]; 
        }

        return $response;

    }

    public function getUbicacionesHistorico($provider_id, $fecha_consulta, $fecha_limite = null){
        $direcciones = [];
        if(is_null($fecha_limite)){
            $ubicaciones_json = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
            WHERE modelo = 'Ubicacion' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'type_address_prov' is null
            AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta ORDER BY fecha_validacion  DESC",
            ['provider_id' => $provider_id, 'tipo' => 'bys', 'fecha_consulta' => $fecha_consulta])->queryAll();
        }else{
            $ubicaciones_json = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                WHERE modelo = 'Ubicacion' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'type_address_prov' is null
                AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta AND fecha_validacion::date > :fecha_limite ORDER BY fecha_validacion  DESC",
                ['provider_id' => $provider_id, 'tipo' => 'bys', 'fecha_consulta' => $fecha_consulta, 'fecha_limite' => $fecha_limite ])->queryAll();
        }
        if( count($ubicaciones_json) > 0 ){
            foreach($ubicaciones_json as $registro){
                if( !empty($registro['item_object']) ){
                    $response = ["ubicacion" => new Ubicacion(), "fotografia" => new FotografiaNegocio()];
                    $response["ubicacion"] = self::setAtributesModel("app\models\\Ubicacion", json_decode($registro['item_object']));
                    if(is_null($fecha_limite)){
                        $fotografia_json = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                            WHERE modelo = 'FotografiaNegocio' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'ubicacion_id' = :ubicacion_id
                            AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta ORDER BY fecha_validacion DESC",
                            ['provider_id' => $provider_id, 'tipo' => 'bys', 'ubicacion_id' => $response["ubicacion"]->ubicacion_id, 'fecha_consulta' => $fecha_consulta])->queryOne();
                    }else{
                        $fotografia_json = Yii::$app->db->createCommand("SELECT arr.item_object FROM provider.historico, json_array_elements(data) WITH ordinality arr(item_object)
                            WHERE modelo = 'FotografiaNegocio' AND provider_id = :provider_id AND tipo = :tipo AND arr.item_object::jsonb->>'ubicacion_id' = :ubicacion_id
                            AND json_typeof(data) = 'array' AND fecha_validacion::date <= :fecha_consulta AND fecha_validacion::date > :fecha_limite ORDER BY fecha_validacion DESC",
                            ['provider_id' => $provider_id, 'tipo' => 'bys', 'ubicacion_id' => $response["ubicacion"]->ubicacion_id, 'fecha_consulta' => $fecha_consulta, 'fecha_limite' => $fecha_limite])->queryOne();
                    }
                    if( !empty($fotografia_json['item_object']) ){ $response["fotografia"] = self::setAtributesModel("app\models\\FotografiaNegocio", json_decode($fotografia_json['item_object'])); }
                    array_push($direcciones, $response);
                }
            }
        }else{
            if(is_null($fecha_limite)){
                $model = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'Ubicacion' AND provider_id = $provider_id AND tipo = 'bys' 
                AND data::json->>'type_address_prov' is null AND json_typeof(data) = 'object' AND fecha_validacion::date <= '$fecha_consulta'")
                ->orderBy(['fecha_validacion' => SORT_DESC])->all();
            }else{
                $model = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'Ubicacion' AND provider_id = $provider_id AND tipo = 'bys' 
                AND data::json->>'type_address_prov' is null AND json_typeof(data) = 'object' AND fecha_validacion::date <= '$fecha_consulta' AND fecha_validacion::date > '$fecha_limite'")
                ->orderBy(['fecha_validacion' => SORT_DESC])->all();
            }
            if( count($model) > 0 ){
                foreach($model as $historico){
                    $response = ["ubicacion" => new Ubicacion(), "fotografia" => new FotografiaNegocio()];
                    $response["ubicacion"] = self::setAtributesModel("app\models\\Ubicacion", $historico->data);
                    if(is_null($fecha_limite)){
                        $histFotografia = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'FotografiaNegocio' AND provider_id = $provider_id AND tipo = 'bys' 
                            AND data::json->>'ubicacion_id' = :ubicacion_id AND json_typeof(data) = 'object' AND fecha_validacion::date <= :fecha_consulta", 
                            ['ubicacion_id' => $response["ubicacion"]->ubicacion_id, 'fecha_consulta' => $fecha_consulta])
                            ->orderBy(['fecha_validacion' => SORT_DESC])->one(); //Busca datos legacy primeramente en caso de migracion
                    } else {
                        $histFotografia = Historico::findBySql("SELECT * FROM provider.historico WHERE modelo = 'FotografiaNegocio' AND provider_id = $provider_id AND tipo = 'bys' 
                        AND data::json->>'ubicacion_id' = :ubicacion_id AND json_typeof(data) = 'object' AND fecha_validacion::date <= :fecha_consulta AND fecha_validacion::date > :fecha_limite", 
                        ['ubicacion_id' => $response["ubicacion"]->ubicacion_id, 'fecha_consulta' => $fecha_consulta, 'fecha_limite' => $fecha_limite])
                        ->orderBy(['fecha_validacion' => SORT_DESC])->one(); //Busca datos legacy primeramente en caso de migracion
                    }
                    if( !empty($histFotografia->data) ){ $response["fotografia"] = self::setAtributesModel("app\models\\FotografiaNegocio", $histFotografia->data); }
                    array_push($direcciones, $response);
                }
            }
        }
        return $direcciones;
    }

    private function setAtributesModel($clase, $dataHistorico){
        $response = new $clase;
        foreach($dataHistorico as $atributo => $valor){
            if($response->hasAttribute($atributo)){ $response->$atributo = $valor; }
        }

        return $response;
    }

    public function actionDescargarReporteHistorico($provider_id, $time){

        // $provider = Provider::findOne($provider_id);

        $date = new DateTime();
        $date->setTimestamp($time);
        $date_string = date_format($date, 'Y-m-d');

        /* $filtro_certificados = HistoricoCertificados::find()->select( new Expression('created_at::date as date') )->where( ['and', ['provider_id' => $provider->provider_id], ['tipo' => 'CERTIFICADO'], ['provider_type' => 'bys'], ['<', 'date(created_at)', $date_string] ] )->orderBy(['created_at' => SORT_DESC])->asArray()->all();
        $fechas_periodos = ArrayHelper::getColumn($filtro_certificados,'date');

        $datos_historico = self::getHistoricoData($provider, $date_string, $fechas_periodos); */

        self::generarExcelHistorico($provider_id, $date_string, true);
        
    }

    public function actionDescargarHistorialCertificacion($provider_id){
        $temp_path = 'cache_tmp/historico/';
        !file_exists($temp_path) && mkdir($temp_path,0775,true);
        $provider = Provider::findOne($provider_id);

        $name_zip = "{$temp_path}{$provider->rfc}_historico.zip";

        $certificados_dates = Yii::$app->db->createCommand("SELECT date FROM ( SELECT DISTINCT(created_at::date) as date FROM historico_certificados 
            WHERE provider_id = $provider_id AND tipo = 'CERTIFICADO' AND provider_type = 'bys' ) x ORDER BY date DESC ")->queryAll();
        $fechas = ArrayHelper::getColumn($certificados_dates,'date');

        $zipData = new \ZipArchive();
        $zipData->open($name_zip, \ZIPARCHIVE::CREATE);

        $arr_temp_files = [];

        foreach($fechas as $fecha){
            $archivo_historico = self::generarExcelHistorico($provider_id, $fecha);
            $zipData->addFile("{$temp_path}{$archivo_historico}",$archivo_historico);
            array_push($arr_temp_files,"{$temp_path}{$archivo_historico}");
        }
        $zipData->close();

        foreach ($arr_temp_files as $file) unlink($file);

        return ReportsController::downloadZip($name_zip);

    }

    /**
     * Genera un excel con la informacion del historico del proveedor hasta la fecha solicitada (inicio de los tiempos - X fecha)
     * @param Provider $provider Instancia del modelo Provider
     * @param String $date_string Fecha a consultar (Y-m-d)
     * @param bool $download Indicador que configura la funcion regresar la data del archivo o solamente el path del documento
     * 
     * @return Null|String Dependiendo del valor de $download regresara la data del archivo para descargarse o el path del documento
     * 
     */
    public function generarExcelHistorico($provider_id, $date_string,  $download = false){
        $provider_data = Provider::findOne($provider_id);
        $temp_path = 'cache_tmp/historico/';

        $filtro_certificados = Yii::$app->db->createCommand("SELECT date FROM ( SELECT DISTINCT(created_at::date) as date FROM historico_certificados 
            WHERE provider_id = $provider_data->provider_id AND tipo = 'CERTIFICADO' AND provider_type = 'bys' ) x WHERE date < '$date_string' ORDER BY date DESC ")->queryAll();
        $fechas_periodos = ArrayHelper::getColumn($filtro_certificados,'date');

        $datos_historico = self::getHistoricoData($provider_data, $date_string, $fechas_periodos);

        $certificado_proveedor =  HistoricoCertificados::find()->where( ['and', ['provider_id' => $provider_data->provider_id], ['tipo' => 'CERTIFICADO'], ['provider_type' => 'bys'], ['date(created_at)' => $date_string] ] )->orderBy(['created_at' => SORT_DESC])->one();
        $cartas_protesta = HistoricoCartaProtesta::find()->where(['and', ['provider_id' => $provider_data->provider_id], ['is not', 'firma_bys', null], ['<=', 'date(created_at)', $date_string]])->orderBy(['created_at' => SORT_DESC])->all();

        #region Perfil
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setTitle("Perfil");
        $sheet->setCellValue('A1', 'Perfil');
        $sheet->getStyle('A1')->getFont()->setSize(13);
        
        $sheet->setCellValue('A2', $provider_data->isPFisica() ? 'Nombre completo' : 'Razón social');
        $sheet->setCellValue('A3', 'RFC');
        $sheet->setCellValue('A4', 'Nombre comercial');
        $sheet->setCellValue('A5', 'Pais de origen');
        $sheet->setCellValue('A6', 'Correo');
        $sheet->setCellValue('A7', 'Correo de registro');
        $sheet->setCellValue('A8', 'Telefono');
        $sheet->setCellValue('A9', 'Estratificacion');
        $sheet->setCellValue('A10', 'Constancia de Situación Fiscal');
        $sheet->setCellValue('A11', 'Correo para cotizaciones');
        $sheet->setCellValue('A12', 'Tel. cotizaciones');
        $sheet->setCellValue('A13', 'Extensión');
        $sheet->setCellValue('A14', 'Contacto para cotizaciones');

        $sheet->setCellValue('B2', $provider_data->getNameOrRazonSocial() );
        $sheet->setCellValue('B3', $provider_data->rfc);
        $sheet->setCellValue('B4', !empty($provider_data->name_comercial) ? $provider_data->name_comercial : 'NA' );
        $sheet->setCellValue('B5', $provider_data->getNacionalidad());
        $sheet->setCellValue('B6', $provider_data->email);
        $sheet->setCellValue('B7', $provider_data->usuario->email);
        $sheet->setCellValue('B8', $provider_data->telfono);
        $sheet->setCellValue('B9', $provider_data->estratificacion);
        $sheet->setCellValue('B10', self::setFullUrlFile($datos_historico['rfc_data']->url_rfc));
        $sheet->setCellValue('B11', $datos_historico['perfil']->correo_cot);
        $sheet->setCellValue('B12', $datos_historico['perfil']->telefono_cot);
        $sheet->setCellValue('B13', $datos_historico['perfil']->ext_cot);
        $sheet->setCellValue('B14', $datos_historico['perfil']->contacto_cot);

        $sheet->getStyle('B2:B14')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheet->getStyle('A2:A14')->getFont()->setBold(true);

        // if(true){ //Falta validar si esta informacion estara de cajon o solo sera por alguna condicion
        //     $sheet->setCellValue('A16', 'Comprobante de domicilio');
        //     $sheet->getStyle('A16')->getFont()->setSize(13);
    
        //     $sheet->setCellValue('A17', 'Documento');
        //     $sheet->setCellValue('A18', 'Fecha de vencimiento');
        //     $sheet->setCellValue('A19', 'Motivo por el cual no cuenta con el tipo de comprobante requerido');
    
        //     $sheet->setCellValue('B17', self::setFullUrlFile($datos_historico['comprobante_domicilio']->url_comprobante_domicilio));
        //     $sheet->setCellValue('B18', $datos_historico['comprobante_domicilio']->expiration_date);
        //     $sheet->setCellValue('B19', $datos_historico['comprobante_domicilio']->no_cuenta_tipo_comprobante);

        //     $sheet->getStyle('B17:B19')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        //     $sheet->getStyle('A17:A19')->getFont()->setBold(true);
        // }

        #endregion

        #region Legales
        $sheetL = $spreadsheet->createSheet();
        $sheetL->setTitle("Legales");

        $sheetL->setCellValue('A1', $provider_data->isPFisica() ? 'Datos Legales' : 'Datos del Acta Constitutiva');
        $sheetL->getStyle('A1')->getFont()->setSize(13);


        if( $provider_data->isPFisica() ){
            $sheetL->setCellValue('A2', 'CURP');
            $sheetL->setCellValue('A3', 'Comprobante CURP');
            $sheetL->setCellValue('A4', 'Tipo de identificación');
            $sheetL->setCellValue('A5', 'Identificación oficial');
            $sheetL->setCellValue('A6', 'Vencimiento');

            $sheetL->setCellValue('B2', $datos_historico['curp_data']->curp);
            $sheetL->setCellValue('B3', self::setFullUrlFile($datos_historico['curp_data']->url_curp));
            $sheetL->setCellValue('B4', $datos_historico['idOf_data']->getTipoIdentificacion() );
            $sheetL->setCellValue('B5', self::setFullUrlFile($datos_historico['idOf_data']->url_idoficial));
            $sheetL->setCellValue('B6', $datos_historico['idOf_data']->getFechaVencimiento());

            $sheetL->getStyle('B2:B7')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            $sheetL->getStyle('A2:A7')->getFont()->setBold(true);
        }else{
            $sheetL->setCellValue('A2', 'Número de acta');
            $sheetL->setCellValue('A3', 'Libro');
            $sheetL->setCellValue('A4', 'Foja');
            $sheetL->setCellValue('A5', 'Nombre y número de fedatario público');
            $sheetL->setCellValue('A6', 'Acta Constitutiva notariada');
            $sheetL->setCellValue('A7', 'Fecha de acta');

            $sheetL->setCellValue('B2', $datos_historico['model_acta_constitutiva']->num_acta);
            $sheetL->setCellValue('B3', !empty($datos_historico['model_acta_constitutiva']->libro) ? $datos_historico['model_acta_constitutiva']->libro : 'NA' );
            $sheetL->setCellValue('B4', !empty($datos_historico['model_acta_constitutiva']->foja) ? $datos_historico['model_acta_constitutiva']->foja : 'NA' );
            $sheetL->setCellValue('B5', !empty($datos_historico['model_acta_constitutiva']->federatario_publico) ? $datos_historico['model_acta_constitutiva']->federatario_publico : 'NA' );
            $sheetL->setCellValue('B6', self::setFullUrlFile($datos_historico['model_acta_constitutiva']->documento_acta));
            $sheetL->setCellValue('B7', $datos_historico['model_acta_constitutiva']->fecha_acta);

            $sheetL->getStyle('B2:B7')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            $sheetL->getStyle('A2:A7')->getFont()->setBold(true);

            $sheetL->setCellValue('A9', 'Modificación acta');
            $sheetL->getStyle('A9')->getFont()->setSize(13);

            $index_celda = 10;
            $index_original = $index_celda;

            foreach($datos_historico['mod_acta_data'] as $mod_acta){
                $sheetL->setCellValue("A$index_celda", 'Fecha acta');
                $sheetL->setCellValue("B$index_celda", $mod_acta->fecha_acta);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Número de acta');
                $sheetL->setCellValue("B$index_celda", $mod_acta->nombre_documento);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Descripción');
                $sheetL->setCellValue("B$index_celda", !empty($mod_acta->descripcion_cambios) ? $mod_acta->descripcion_cambios : 'NA');
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Documento acta');
                $sheetL->setCellValue("B$index_celda", self::setFullUrlFile($mod_acta->documento_acta));
                $index_celda += 2;
            }

            $sheetL->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            $sheetL->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true); $index_celda += 2;

            $sheetL->setCellValue("A$index_celda", 'Datos del representante legal');
            $sheetL->getStyle("A$index_celda")->getFont()->setSize(13); $index_celda += 2;

            $index_original = $index_celda;

            foreach($datos_historico['representante_data'] as $rep_legal){
                $sheetL->setCellValue("A$index_celda", 'RFC');
                $sheetL->setCellValue("B$index_celda", $rep_legal->rfc);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'CURP');
                $sheetL->setCellValue("B$index_celda", $rep_legal->curp);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Nombre');
                $sheetL->setCellValue("B$index_celda", $rep_legal->nombre);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Primer apellido');
                $sheetL->setCellValue("B$index_celda", $rep_legal->ap_paterno);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Segundo apellido');
                $sheetL->setCellValue("B$index_celda", $rep_legal->ap_materno);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Teléfono');
                $sheetL->setCellValue("B$index_celda", $rep_legal->telefono);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Tipo de identificación');
                $sheetL->setCellValue("B$index_celda", $rep_legal->getTipoIdentificacion());
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Vencimiento');
                $sheetL->setCellValue("B$index_celda", $rep_legal->getFechaVencimiento());
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Firmante');
                $sheetL->setCellValue("B$index_celda", $rep_legal->rep_bys ? 'Si' : 'No');
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Documento de identificación');
                $sheetL->setCellValue("B$index_celda", self::setFullUrlFile($rep_legal->documento_identificacion));
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", 'Poder Otorgado');
                $sheetL->setCellValue("B$index_celda", $rep_legal->tipo_poder);
                $index_celda++;
                $sheetL->setCellValue("A$index_celda", ucfirst(mb_strtolower($rep_legal->tipo_poder)));
                $sheetL->setCellValue("B$index_celda", self::setFullUrlFile($rep_legal->getDocumentoRelacion()));
                $index_celda += 2;
            }

            $sheetL->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            $sheetL->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true); $index_celda += 2;

            $sheetL->setCellValue("A$index_celda", 'Accionistas');
            $sheetL->getStyle("A$index_celda")->getFont()->setSize(13); $index_celda += 2;

            $index_original = $index_celda;

            foreach($datos_historico['accionistas_data'] as $accionista){
                $sheetL->setCellValue("A$index_celda", 'RFC');
                $sheetL->setCellValue("B$index_celda", $accionista->rfc);
                $index_celda++;
                if( strlen($accionista->rfc) == 12 ){
                    $sheetL->setCellValue("A$index_celda", 'Razón social');
                    $sheetL->setCellValue("B$index_celda", $accionista->razon_social);
                    $index_celda++;
                }else{
                    $sheetL->setCellValue("A$index_celda", 'CURP');
                    $sheetL->setCellValue("B$index_celda", $accionista->curp);
                    $index_celda++;
                    $sheetL->setCellValue("A$index_celda", 'Nombre');
                    $sheetL->setCellValue("B$index_celda", $accionista->nombre);
                    $index_celda++;
                    $sheetL->setCellValue("A$index_celda", 'Primer apellido');
                    $sheetL->setCellValue("B$index_celda", $accionista->ap_paterno);
                    $index_celda++;
                    $sheetL->setCellValue("A$index_celda", 'Segundo apellido');
                    $sheetL->setCellValue("B$index_celda", $accionista->ap_materno);
                    $index_celda++;
                }
                
                $sheetL->setCellValue("A$index_celda", ucfirst(mb_strtolower($accionista->tipo_relacion)));
                $sheetL->setCellValue("B$index_celda", self::setFullUrlFile($accionista->getDocumentoRelacion()));
                $index_celda += 2;
            }

            $sheetL->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            $sheetL->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true);

        }
        #endregion

        #region Actividad Economica
        $sheetAE = $spreadsheet->createSheet();
        $sheetAE->setTitle("Actividad economica");

        $sheetAE->setCellValue('A1', 'Actividad económica');
        $sheetAE->getStyle('A1')->getFont()->setSize(13);

        $index_celda = 3;
        $index_original = $index_celda;

        foreach($datos_historico['actividades_data'] as $actividad){
            $sheetAE->setCellValue("A$index_celda", 'Sector');
            $sheetAE->setCellValue("B$index_celda", $actividad->sectors);
            $index_celda++;
            $sheetAE->setCellValue("A$index_celda", 'Rama');
            $sheetAE->setCellValue("B$index_celda", $actividad->rama);
            $index_celda++;
            $sheetAE->setCellValue("A$index_celda", 'Nombre de la actividad');
            $sheetAE->setCellValue("B$index_celda", isset($actividad->actividad->nombre_actividad)?$actividad->actividad->nombre_actividad:'');
            $index_celda++;
            $sheetAE->setCellValue("A$index_celda", 'Fecha inicio de la actividad económica	');
            $sheetAE->setCellValue("B$index_celda", $actividad->start_date);
            $index_celda++;
            $sheetAE->setCellValue("A$index_celda", 'Porcentaje');
            $sheetAE->setCellValue("B$index_celda", $actividad->porcentaje);
            $index_celda++;
            if( !empty($actividad->url_factura_c_c) ){
                $sheetAE->setCellValue("A$index_celda", 'Factura');
                $sheetAE->setCellValue("B$index_celda", self::setFullUrlFile($actividad->url_factura_c_c));
                $index_celda++;
            }
            $index_celda++;
        }

        $sheetAE->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetAE->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true); $index_celda += 2;

        $sheetAE->setCellValue("A$index_celda", 'Productos');
        $sheetAE->getStyle("A$index_celda")->getFont()->setSize(13); $index_celda += 2;

        $index_original = $index_celda;
        foreach($datos_historico['productos_data'] as $producto){
            if( $producto["mode"] == 1 ){
                $sheetAE->setCellValue("A$index_celda", 'Familia');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getFamiliaV2($producto['grupo']));$index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Grupo');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getGrupoV2($producto['grupo'])); $index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Linea');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getLineaV2($producto['producto_id'])); $index_celda += 2;
            } else if( $producto["mode"] == 2 ){
                $sheetAE->setCellValue("A$index_celda", 'Grupo');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getGrupoV1($producto['grupo'])); $index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Producto');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getNombreProductoV1($producto['producto_id'])); $index_celda += 2;
            }else{
                $sheetAE->setCellValue("A$index_celda", 'Familia');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getFamiliaV3($producto['familia']));$index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Grupo');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getGrupoV3($producto['grupo'])); $index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Clase');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getClaseV3($producto['clase'])); $index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Linea');
                $sheetAE->setCellValue("B$index_celda", ProviderGiro::getProductosV3($producto['producto_id'])); $index_celda++;
                $sheetAE->setCellValue("A$index_celda", 'Factura');
                $sheetAE->setCellValue("B$index_celda", self::setFullUrlFile(!empty($producto['factura'])? $producto['factura'] : null));
                $index_celda += 2;
            }
        }

        $sheetAE->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetAE->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true);
        #endregion

        #region Capacidad Economica
        $sheetCE = $spreadsheet->createSheet();
        $sheetCE->setTitle("Capacidad Económica");

        $sheetCE->setCellValue('A1', 'Capacidad Económica');
        $sheetCE->getStyle('A1')->getFont()->setSize(13);

        $sheetCE->setCellValue('A2', 'Declaración del ejercicio fiscal más reciente');
        $sheetCE->setCellValue('A3', 'Año de la Declaración Anual');
        $sheetCE->setCellValue('A4', 'Opinión de cumplimiento positiva por el SAT');
        $sheetCE->setCellValue('A5', 'Pago o justificante por no pagar ISN');
        $sheetCE->setCellValue('A6', 'Opinión de cumplimiento por la tesorería del estado');
        $sheetCE->setCellValue('A7', 'Balance general año en curso');
        $sheetCE->setCellValue('A8', 'Estado de resultados año en curso');
        $sheetCE->setCellValue('A9', 'Balance general año anterior');
        $sheetCE->setCellValue('A10', 'Estado de resultados año anterior');
        $sheetCE->setCellValue('A11', 'Comprobante de pago');
        $sheetCE->setCellValue('A12', 'Acuse de la última declaración de pago provisional');
        $sheetCE->setCellValue('A13', 'Ventas Anuales de la Declaración Anual');
        $sheetCE->setCellValue('A14', 'Número de personas laborando');
        $sheetCE->setCellValue('A15', 'Estratificación');

        $sheetCE->setCellValue('B2', self::setFullUrlFile($datos_historico['ult_declaracion_data']->declaracion_ejercicio_fiscal));
        $sheetCE->setCellValue('B3', $datos_historico['ult_declaracion_data']->ejercicio_presentado);
        $sheetCE->setCellValue('B4', self::setFullUrlFile($datos_historico['ult_declaracion_data']->comprobante_pago));
        $sheetCE->setCellValue('B5', self::setFullUrlFile($datos_historico['ult_declaracion_data']->justificante_isn));
        $sheetCE->setCellValue('B6', self::setFullUrlFile($datos_historico['ult_declaracion_data']->cumplimiento_estado));
        $sheetCE->setCellValue('B7', self::setFullUrlFile($datos_historico['ult_declaracion_data']->balance_general));
        $sheetCE->setCellValue('B8', self::setFullUrlFile($datos_historico['ult_declaracion_data']->estado_resultado));
        $sheetCE->setCellValue('B9', self::setFullUrlFile($datos_historico['ult_declaracion_data']->balance_general_anterior));
        $sheetCE->setCellValue('B10', self::setFullUrlFile($datos_historico['ult_declaracion_data']->estado_resultado_anterior));
        $sheetCE->setCellValue('B11', self::setFullUrlFile($datos_historico['ult_declaracion_data']->comprobantes_de_pago));
        $sheetCE->setCellValue('B12', self::setFullUrlFile($datos_historico['ult_declaracion_data']->ultima_declaracion_pago_prov));
        $sheetCE->setCellValue('B13', $datos_historico['ult_declaracion_data']->ingresos_mercantiles);
        $sheetCE->setCellValue('B14', $datos_historico['ult_declaracion_data']->personal);
        $sheetCE->setCellValue('B15', $datos_historico['provider_ext']->estratificacion);

        $sheetCE->getStyle('B2:B15')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetCE->getStyle('A2:A15')->getFont()->setBold(true);

        #endregion

        #region Experiencia Comercial
        $sheetEC = $spreadsheet->createSheet();
        $sheetEC->setTitle("Experiencia comercial");

        $sheetEC->setCellValue('A1', 'Experiencia comercial');
        $sheetEC->getStyle('A1')->getFont()->setSize(13);

        $sheetEC->setCellValue('A2', 'Protocolo');
        $sheetEC->setCellValue('A3', 'Página web');
        $sheetEC->setCellValue('A4', 'Currículum');

        $sheetEC->setCellValue('B2', $provider_data->protocolo);
        $sheetEC->setCellValue('B3', $provider_data->pagina_web);
        $sheetEC->setCellValue('B4', self::setFullUrlFile($provider_data->getCurriculumValue()));

        $sheetEC->getStyle("B2:B4")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetEC->getStyle("A2:A4")->getFont()->setBold(true);

        $sheetEC->setCellValue('A6', 'Clientes');
        $sheetEC->getStyle('A6')->getFont()->setSize(13);

        $index_celda = 8;
        $index_original = $index_celda;
        foreach($datos_historico['clientes_data'] as $cliente_data){
            $sheetEC->setCellValue("A$index_celda", 'Nombre o razón social');
            $sheetEC->setCellValue("B$index_celda", $cliente_data->nombre_razon_social);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Contacto');
            $sheetEC->setCellValue("B$index_celda", $cliente_data->persona_dirigirse);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'RFC');
            $sheetEC->setCellValue("B$index_celda", $cliente_data->rfc);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Teléfono');
            $sheetEC->setCellValue("B$index_celda", $cliente_data->telefono);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Factura');
            $sheetEC->setCellValue("B$index_celda", !empty($cliente_data->url_factura_c_c) ? self::setFullUrlFile($cliente_data->url_factura_c_c) : 'NA');
            $index_celda += 2;
        }

        $sheetEC->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetEC->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true);

        $sheetEC->setCellValue("A$index_celda", 'Certificados');
        $sheetEC->getStyle("A$index_celda")->getFont()->setSize(13); $index_celda += 2;

        $index_original = $index_celda;

        foreach($datos_historico['certificados_data'] as $certificado){
            $sheetEC->setCellValue("A$index_celda", 'Nombre del documento');
            $sheetEC->setCellValue("B$index_celda", $certificado->certificador);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Indefinido');
            $sheetEC->setCellValue("B$index_celda", $certificado->undefined ? 'Sí' : 'No');
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Vigencia');
            $sheetEC->setCellValue("B$index_celda", $certificado->vigencia);
            $index_celda++;
            $sheetEC->setCellValue("A$index_celda", 'Documento');
            $sheetEC->setCellValue("B$index_celda", self::setFullUrlFile($certificado->url_archivo));
            $index_celda += 2;
        }

        $sheetEC->getStyle("B$index_original:B$index_celda")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetEC->getStyle("A$index_original:A$index_celda")->getFont()->setBold(true);
        #endregion

        #region Ubicacion
        $sheetUb = $spreadsheet->createSheet();
        $sheetUb->setTitle("Ubicacion");

        $sheetUb->setCellValue('A1', 'Domicilio Fiscal');
        $sheetUb->getStyle('A1')->getFont()->setSize(13); 

        $index_ub = 2;

        $returnIndex = self::setUbicacionExcel($sheetUb, $datos_historico['ubicacion_fiscal_data'], $index_ub); $returnIndex += 2;
        if( isset( $datos_historico['ubicacion_nl_data']['ubicacion'] ) && !empty( $datos_historico['ubicacion_nl_data']['ubicacion'] ) ) { 
            $sheetUb->setCellValue("A$returnIndex", 'Ubicacion Nuevo Leon');
            $sheetUb->getStyle("A$returnIndex")->getFont()->setSize(13); $returnIndex++;
            $returnIndex = self::setUbicacionExcel($sheetUb, $datos_historico['ubicacion_nl_data'], $returnIndex); $returnIndex += 2;
        }

        if( count($datos_historico['ubicaciones_data']) != 0){
            $sheetUb->setCellValue("A$returnIndex", 'Ubicaciones');
            $sheetUb->getStyle("A$returnIndex")->getFont()->setSize(13); $returnIndex++;
            foreach($datos_historico['ubicaciones_data'] as $ubicaconMin){
                $returnIndex = self::setUbicacionExcel($sheetUb, $ubicaconMin, $returnIndex);
            }
        }
        
        #endregion

        #region Bancos
        $sheetDB = $spreadsheet->createSheet();
        $sheetDB->setTitle("Datos Bancarios");

        $sheetDB->setCellValue('A1', 'Datos Bancarios');
        $sheetDB->getStyle('A1')->getFont()->setSize(13);

        $sheetDB->setCellValue('A2', 'Nombre del titular de la cuenta	');
        $sheetDB->setCellValue('A3', 'Banco');
        $sheetDB->setCellValue('A4', 'Clabe');
        $sheetDB->setCellValue('A5', 'Estado de cuenta');

        $sheetDB->setCellValue('B2', $datos_historico['pago_data']->nombre_titular_cuenta);
        $sheetDB->setCellValue('B3', isset($datos_historico['pago_data']->bancos->banco)?$datos_historico['pago_data']->bancos->banco:'');
        $sheetDB->setCellValue('B4', $datos_historico['pago_data']->cuenta_clave);
        $sheetDB->setCellValue('B5', self::setFullUrlFile($datos_historico['pago_data']->estado_cuenta));

        $sheetDB->getStyle('B2:B5')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetDB->getStyle('A2:A5')->getFont()->setBold(true);
        #endregion

        #region Carta protesta
        $sheetCP = $spreadsheet->createSheet();
        $sheetCP->setTitle("Carta protesta");

        $sheetCP->setCellValue('A1', 'Carta protesta');
        $sheetCP->getStyle('A1')->getFont()->setSize(13);

        $index_celda_cp = 2;
        $index_original_cp = $index_celda_cp;

        foreach($cartas_protesta as $carta_protesta){
            $sheetCP->setCellValue("A$index_celda_cp", 'Fecha');
            $sheetCP->setCellValue("B$index_celda_cp", $carta_protesta->created_at);
            $index_celda_cp++;
            $sheetCP->setCellValue("A$index_celda_cp", 'Documento');
            $sheetCP->setCellValue("B$index_celda_cp", self::setFullUrlFile($carta_protesta->url_carta));                
            $index_celda_cp += 2;
        }

        $sheetCP->getStyle("B$index_original_cp:B$index_celda_cp")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetCP->getStyle("A$index_original_cp:A$index_celda_cp")->getFont()->setBold(true);
        #endregion

        #region Certificado
        $sheetCer = $spreadsheet->createSheet();
        $sheetCer->setTitle("Certificado");

        $sheetCer->setCellValue('A1', 'Certificado');
        $sheetCer->getStyle('A1')->getFont()->setSize(13);

        $sheetCer->setCellValue('A2', 'Fecha');
        $sheetCer->setCellValue('A3', 'Documento');

        $sheetCer->setCellValue('B2', $date_string);
        $sheetCer->setCellValue('B3', self::setFullUrlFile($certificado_proveedor->url_certificado));

        $sheetCer->getStyle('B2:B3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetCer->getStyle('A2:A3')->getFont()->setBold(true);
        #endregion

        #region Curso

        if(true){ /* Se habilitara solo cuando lo pidan */

            $curso_proveedor =  HistoricoCertificados::find()->where( ['and', ['provider_id' => $provider_data->provider_id], ['tipo' => 'CURSO'], ['provider_type' => 'bys'], ['<=', 'date(created_at)', $date_string] ] )->orderBy(['created_at' => SORT_DESC])->one();

            $sheetCur = $spreadsheet->createSheet();
            $sheetCur->setTitle("Curso");

            $sheetCur->setCellValue('A1', 'Curso');
            $sheetCur->getStyle('A1')->getFont()->setSize(13);

            if( !is_null($curso_proveedor) ){
                $sheetCur->setCellValue('A2', 'Fecha');
                $sheetCur->setCellValue('A3', 'Documento');

                $sheetCur->setCellValue('B2', $date_string);
                $sheetCur->setCellValue('B3', self::setFullUrlFile($curso_proveedor->url_certificado));

                $sheetCur->setCellValue('B2', !is_null($curso_proveedor) ? $curso_proveedor->created_at : 'NA');
                $sheetCur->setCellValue('B3', !is_null($curso_proveedor) ? self::setFullUrlFile($curso_proveedor->url_certificado) : 'NA');

                $sheetCur->getStyle('B2:B3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $sheetCur->getStyle('A2:A3')->getFont()->setBold(true);
            }

        }
        
        #endregion

        #region Preregistro

        $preregistro = HistoricoCertificados::find()->where(['and', ['tipo' => 'PREREGISTRO', 'provider_id' => $provider_data->provider_id]])->orderBy(['created_at' => SORT_DESC])->one();

        $sheetPre = $spreadsheet->createSheet();
        $sheetPre->setTitle("Pre-Registro");

        $sheetPre->setCellValue('A1', 'Pre registro');
        $sheetPre->getStyle('A1')->getFont()->setSize(13);

        $sheetPre->setCellValue('A2', 'Fecha');
        $sheetPre->setCellValue('A3', 'Documento');

        $sheetPre->setCellValue('B2', !is_null($preregistro) ? $preregistro->created_at : 'NA');
        $sheetPre->setCellValue('B3', !is_null($preregistro) ? self::setFullUrlFile($preregistro->url_certificado) : 'NA');

        $sheetPre->getStyle('B2:B3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetPre->getStyle('A2:A3')->getFont()->setBold(true);

        #endregion

        #region Visita

        //A quien vea esto, la neta no me la quise complicar tanto
        $visitas = Yii::$app->db->createCommand("SELECT vd.visit_id, vd.ubicacion_id FROM visit_details vd INNER JOIN
        visit v on v.visit_id = vd.visit_id WHERE v.provider_id = $provider_data->provider_id AND date(v.created_at) <= '{$date_string}' ORDER BY v.created_at")->queryAll();

        $arr_visitas = ArrayHelper::map($visitas,'visit_id', 'ubicacion_id');

        $sheetVisit = $spreadsheet->createSheet();
        $sheetVisit->setTitle("Visitas");

        $sheetVisit->setCellValue('A1', 'Documento(s) de visitas');
        $sheetVisit->getStyle('A1')->getFont()->setSize(13);

        $index_celda_visita = 3;

        foreach($arr_visitas as $visita => $ubicacion){
            $sheetVisit->setCellValue("A$index_celda_visita", self::setFullUrlFile("visit/formato?v={$visita}&u={$ubicacion}"));
            $index_celda_visita++;
        }

        #endregion

        #region Cotejo
        $cotejos = DatosValidados::find()->where(['and', ['modelo' => 'cotejar', 'provider_id' => $provider_data->provider_id], ['<=', 'date(created_date)', $date_string]])->orderBy(['created_date' => SORT_DESC])->all();

        $sheetPre = $spreadsheet->createSheet();
        $sheetPre->setTitle("Cotejo");

        $sheetPre->setCellValue('A1', 'Acta cotejo');
        $sheetPre->getStyle('A1')->getFont()->setSize(13);

        $index_celda_cotejo = 2;
        $index_original_cotejo = $index_celda_cotejo;

        foreach($cotejos as $cotejo){
            $sheetPre->setCellValue("A$index_celda_cotejo", 'Fecha');
            $sheetPre->setCellValue("B$index_celda_cotejo", $cotejo->created_date);
            $index_celda_cotejo++;
            $sheetPre->setCellValue("A$index_celda_cotejo", 'Documento');
            $sheetPre->setCellValue("B$index_celda_cotejo", self::setFullUrlFile($cotejo->file));                
            $index_celda_cotejo += 2;
        }

        $sheetPre->getStyle("B$index_original_cotejo:B$index_celda_cotejo")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetPre->getStyle("A$index_original_cotejo:A$index_celda_cotejo")->getFont()->setBold(true);
        #endregion

        #region Sanciones
        $sanciones =  Sanciones::find()->where(['and', ['activo' => 1, 'proveedor_id' => $provider_data->provider_id]])->all();

        $sheetSan = $spreadsheet->createSheet();
        $sheetSan->setTitle("Sanciones");

        $sheetSan->setCellValue('A1', 'Sanciones');
        $sheetSan->getStyle('A1')->getFont()->setSize(13);

        $index_celda_sanciones = 2;
        $index_original_sanciones = $index_celda_sanciones;

        foreach($sanciones as $sancion){

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Nombre, denominación o razón social');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->provider->getFullName());
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Expediente');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->expediente);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Documentos de la Sanción');
            $sheetSan->setCellValue("B$index_celda_sanciones", self::setFullUrlFile($sancion->documento_sancion));
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Fecha de la resolución sancionadora');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->f_resolucion);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Fecha de notificación de la resolución sancionadora al proveedor o participante');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->f_notificacion);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Fecha de inscripción de la sanción en el Registro de Proveedores y Participantes Sancionados');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->inscripcion);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Sanción y monto');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->sancion_monto);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Vigencia de la publicación de la sanción');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->vigencia);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Fecha de Finalización de la sanción');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->f_fin);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Días faltantes para cumplir la sanción');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->getDias());
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Autoridad sancionadora');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->autoridad_sancionadora);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Motivo');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->motivo);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Comentario');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->comentario);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Status');
            $sheetSan->setCellValue("B$index_celda_sanciones", $sancion->status);
            $index_celda_sanciones++;

            $sheetSan->setCellValue("A$index_celda_sanciones", 'Tipo');
            $sheetSan->setCellValue("B$index_celda_sanciones", Sanciones::getTipos()[$sancion->tipo]);
            $index_celda_sanciones+=2;

        }

        $sheetSan->getStyle("B$index_original_sanciones:B$index_celda_sanciones")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetSan->getStyle("A$index_original_sanciones:A$index_celda_sanciones")->getFont()->setBold(true);

        #endregion Sanciones

        #region Contrataciones

        $contrataciones = HistoricoContrataciones::findAll(['provider_id' => $provider_data->provider_id]);

        $sheetContrataciones = $spreadsheet->createSheet();
        $sheetContrataciones->setTitle("Contrataciones");

        $sheetContrataciones->setCellValue('A1', 'Historico de Contrataciones');
        $sheetContrataciones->getStyle('A1')->getFont()->setSize(13);

        $index_celda_contrataciones = 2;
        $index_original_contrataciones = $index_celda_contrataciones;

        foreach($contrataciones as $h_contrato){
            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", '# Proveedor');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->provider->Clave_ProveedorSire);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Nombre/RazonSocial');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->provider->getFullName());
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'RFC');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->provider->getRfc());
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Vigencia');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->provider->vigencia);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Num Requisicion');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->num_requisicion);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Fecha Requisicion');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->fecha_requisicion);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Orden Compra');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->orden_compra);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Fecha Ord. Compra');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->fecha_orden_compra);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Num Contrato');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->num_contrato);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Obj Adjudicacion');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->obj_adjudicacion);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Cumplimiento');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->cumplimiento);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Fecha Notificacion');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->fecha_notificacion);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Fecha Registro Incumplimiento');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", $h_contrato->fecha_registro_incumplimiento);
            $index_celda_contrataciones++;

            $sheetContrataciones->setCellValue("A$index_celda_contrataciones", 'Documento');
            $sheetContrataciones->setCellValue("B$index_celda_contrataciones", self::setFullUrlFile($h_contrato->documento));
            $index_celda_contrataciones+=2;
        }

        $sheetContrataciones->getStyle("B$index_original_contrataciones:B$index_celda_contrataciones")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheetContrataciones->getStyle("A$index_original_contrataciones:A$index_celda_contrataciones")->getFont()->setBold(true);

        #endregion Contrataciones
        
        $name_file = "{$provider_data->rfc}_{$date_string}.xlsx";

        $writer = new Xlsx($spreadsheet);
        if($download){
            ob_clean();
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="'. urlencode($name_file).'"');
            $writer->save('php://output');
            exit();
            return null;
        }else{
            $file_path = "{$temp_path}{$name_file}";
            $writer->save($file_path);
            return $name_file;
        }
        
    }

    public function setUbicacionExcel($sheet, $ubicacion, $index){
        $newIndex = $index;
        $sheet->setCellValue("A$newIndex", 'Nombre del contacto');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->encargado); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Correo electrónico');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->correo); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Departamento');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->departamento); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Tipo');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->tipo); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Calle');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->calle_fiscal); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'No. Exterior');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->num_ext_fiscal); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'No. Interior');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->num_int_fiscal); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Colonia');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->getNameColonia()); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Código Postal');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->cp_fiscal); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Localidad');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->getNameLocalidad()); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Ciudad');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->getNameCiudad()); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Estado');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->getNameEstado()); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Teléfono');
        $sheet->setCellValue("B$newIndex", $ubicacion['ubicacion']->telefono); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Comprobante domicilio');
        $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['ubicacion']->url_comprobante_domicilio)); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Video 1');
        $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['ubicacion']->url_video)); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Video 2');
        $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['ubicacion']->url_video_2)); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Video 3');
        $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['ubicacion']->url_video_3)); $newIndex++;
        $sheet->setCellValue("A$newIndex", 'Escrito');
        $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['ubicacion']->url_escrito));

        $sheet->getStyle("B$index:B$newIndex")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheet->getStyle("A$index:A$newIndex")->getFont()->setBold(true);

        if( !is_null($ubicacion['fotografia']) ) { 
            $newIndex += 2;

            $sheet->setCellValue("A$newIndex", 'Fotografia(s)');
            $sheet->getStyle("A$newIndex")->getFont()->setSize(13); 

            $newIndex += 2;

            if( $ubicacion['fotografia'] instanceof FotografiaNegocio ){
                $sheet->setCellValue("A$newIndex", 'Archivo');
                $sheet->setCellValue("B$newIndex", self::setFullUrlFile($ubicacion['fotografia']->url_archivo));$newIndex++;
                $sheet->setCellValue("A$newIndex", 'Descripción');
                $sheet->setCellValue("B$newIndex", $ubicacion['fotografia']->descripcion);$newIndex++;
            }else{
                foreach($ubicacion['fotografia'] as $data_fotografia){
                    $sheet->setCellValue("A$newIndex", 'Archivo');
                    $sheet->setCellValue("B$newIndex", self::setFullUrlFile($data_fotografia->url_archivo));$newIndex++;
                    $sheet->setCellValue("A$newIndex", 'Descripción');
                    $sheet->setCellValue("B$newIndex", $data_fotografia->descripcion);$newIndex += 2;;
                }
            }
        }

        $sheet->getStyle("B$index:B$newIndex")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheet->getStyle("A$index:A$newIndex")->getFont()->setBold(true);

        return $newIndex;

    }

    public function setFullUrlFile($value){
        if( is_null($value) || empty($value) ){ return null; }
        return (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] != "off") ? "https://".$_SERVER['SERVER_NAME']."/$value" : "http://".$_SERVER['SERVER_NAME']."/$value";
    }


}
