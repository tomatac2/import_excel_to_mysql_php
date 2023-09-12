

    1- html upload file form  
     
    <!DOCTYPE html>
    <html>
    <body>
    
    <form method="POST" enctype="multipart/form-data">
      Select excel to upload:
      <input type="file" name="excle_sheet" >
      <input type="hidden" name="_csrfToken"  value="<?=$this->request->getAttribute('csrfToken');?>">
      <input type="submit" value="Upload Excel Sheet" name="submit">
    </form>
    
    </body>
    </html>
    
    
    
    ===================================
    2- install package by composer 
    ===================================
    
    composer require phpoffice/phpspreadsheet
    
    
    
    
    =============================
    3- load package in controller 
    ==============================
    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    
    
    
    
    =========================
    4- php code 
    =========================
    
    $excelMimes = array('text/xls', 'text/xlsx', 'application/excel', 'application/vnd.msexcel', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); 
      // Validate whether selected file is a Excel file 
        if( in_array($_FILES['excle_sheet']['type'], $excelMimes)){ 
            // If the file is uploaded 
            if(is_uploaded_file($_FILES['excle_sheet']['tmp_name'])){ 
              $reader = new Xlsx(); 
              $spreadsheet = $reader->load($_FILES['excle_sheet']['tmp_name']); 
              $worksheet = $spreadsheet->getActiveSheet();  
              $worksheet_arr = $worksheet->toArray(); 
      
          // Remove header row 
          unset($worksheet_arr[0]); 
  
          foreach($worksheet_arr as $row){ 
              if($row[0]):
              $student_name = $row[0]; 
              $gender = $row[1]; 
              $national_id = $row[2]; 
              $address = $row[3]; 
              $hajri_date = $row[4]; 
              $day = $row[5]; 
              $hours = $row[6]; 
              $teacher_name = $row[7]; 
              
              $fields = [
                  "student_name"=>$student_name,
                  "gender"=>$gender,
                  "national_id"=>$national_id,
                  "address"=>$course_title,
                  "hajri_date"=>$hajri_date,
                  "day"=>$day,
                  "hours"=>$hours,
                  "teacher_name"=>$teacher_name,
              ];
  
              //save to DB 
              $saveExcle = $this->Certificate->newEmptyEntity();
              $saveExcle = $this->Certificate->patchEntity($saveExcle , $fields);
              $this->Certificate->save($saveExcle);
          endif;
      }


  }
  }



