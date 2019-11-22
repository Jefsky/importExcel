<?php
/**
     * importExcel
     * @return mixed 
     */
    public function importExcel(){
        header("content-type:text/html;charset=utf-8");
        //上传excel文件
        $file = request()->file('excelfile');
        if(empty($file)) $this->error('上传文件为空');
        $org_name = $file->getInfo('name');
        $extension = pathinfo($org_name)['extension'];
        $batch = Db::table('')->where(['filename'=>explode('.', $org_name)[0]])->find();
        if($batch)$this->error('已上传该文件(文件名重复)'); 
        if(!in_array($extension,['xls','xlsx']))$this->error('文件格式错误');
        //文件锁--防止多次上传
        $fp_lock = RUNTIME_PATH.'import_lock.txt';
        $lock_file = fopen($fp_lock, "w+");
        if(flock($lock_file,LOCK_EX|LOCK_NB)){
          ignore_user_abort(true); //客户端关闭仍执行
          //移到/public/uploads/excel/下
          $info = $file->move(ROOT_PATH.'public'.DS.'upload'.DS.'excel');
          //上传文件成功
          if ($info) {
            //引入PHPExcel类
            vendor('PHPExcel.PHPExcel.Reader');
            //获取上传后的文件名
            $fileName = $info->getSaveName();
            //文件路径
            $filePath = ROOT_PATH.'public'.DS.'upload'.DS.'excel'.DS.$fileName;

            //导入批次记录
            $batch_data['filepath'] = 'excel'.DS.$fileName;
            $batch_data['create_time'] = time();
            $batch_data['filename'] = $org_name;
            $batch_data['savename'] = $fileName;
            $data_count = 0;//数量统计            
            $batch_data['data_count'] = $data_count;
            $batch_id = Db::table('')->insertGetId($batch_data);

            //实例化PHPExcel类
            if($extension == 'xlsx'){
              $PHPReader = new \PHPExcel_Reader_Excel2007();
            }else{
              $PHPReader = new \PHPExcel_Reader_Excel5();
            }
            //读取excel文件
            $objPHPExcel = $PHPReader->load($filePath);
            //读取excel文件中的第一个工作表
            $sheet = $objPHPExcel->getSheet(0);
            $allRow = $sheet->getHighestRow();  //取得总行数
            //$allColumn = $sheet->getHighestColumn();  //取得总列数
            //从第二行开始插入，第一行是列名
            $head_office = $top_keyword = $keyword = '';
            for ($j=3; $j <= $allRow; $j++) {
              $data['enterprise'] = $objPHPExcel->getActiveSheet()->getCell("A".$j)->getValue();
              $data['record_name'] = $objPHPExcel->getActiveSheet()->getCell("B".$j)->getValue();
              $data['pop_name'] = $objPHPExcel->getActiveSheet()->getCell("C".$j)->getValue();
              $data['area'] = $objPHPExcel->getActiveSheet()->getCell("D".$j)->getValue();
              $data['head_office'] = $head_office;
              $data['batch_id'] = $batch_id;
              $keyword_arr = [];
              $keyword = $objPHPExcel->getActiveSheet()->getCell("E".$j)->getValue();
              $keyword_arr[] = $data['enterprise'];
              $keyword_arr[] = $top_keyword;
              $keyword_arr[] = $keyword;
              $data['keyword'] = implode(',', array_filter($keyword_arr));
              if(empty($data['record_name']) && empty($data['pop_name']) & empty($data['area'])){
                $head_office = $data['enterprise'];
                $top_keyword = $objPHPExcel->getActiveSheet()->getCell("E".$j)->getValue();
                continue;
              }
              foreach ($data as $key => $value) {
                  // $data[$key] = trim($value);
                  $data[$key] = str_replace(' ', '', $value);
              }
              $data['create_time'] = $data['update_time'] = time();
              $last_id = Db::table('')->insertGetId($data);//保存数据，并返回主键id
              $data_count ++;
            }
            Db::table('')->where(['id'=>$batch_id])->update(['data_count'=>$data_count]);
            //更新公司表
            model('')->updateCompany();
            fclose($lock_file);
            $this->success('数据导入成功');
          }else{
              fclose($lock_file);
              $this->error('数据导入失败');
          }          
        }else{
          $this->error('正在上传,请勿重复操作!');
        }

      }