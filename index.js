var Excel = require('exceljs');
var wb = new Excel.Workbook();
var path = require('path');
var filePath = path.resolve(__dirname,'product.xlsx');
//console.log(filePath)
var mysql = require('mysql2');
// const { isEmpty } = require('newrelic/lib/util/properties');
const connection = mysql.createConnection({
host : process.env.HOST,
user : process.env.USER,
password : process.env.PASSWORD,
database : process.env.DATABASE
})
connection.connect(function(err){
    if(err) throw err ;
    else{
        console.log("db established successfully")
    }
})
const xlsx = require("xlsx");
const fs = require("fs");


let filenames = fs.readdirSync(__dirname);
console.log(filenames);
  
console.log("\nFilenames in directory:");
filenames.forEach((file) => {
    if(file == 'product.xlsx'){
    const _file = path.resolve(__dirname, file);
    console.log(_file);

    var workbook = xlsx.readFile(_file);
    if(workbook.SheetNames!='Sheet1')
    {   
         
   workbook.SheetNames.forEach((wk)=>{
wb.xlsx.readFile(filePath).then(function(){
    let count_ins =0;
    let count_upt =0;
    var sh = wb.getWorksheet(wk);
    for (let i = 2; i <= sh.rowCount; i++) {

      
        
        if(sh.getRow(i).getCell(1).value != null && sh.getRow(i).getCell(1).value != undefined && sh.getRow(i).getCell(2).value != null && sh.getRow(i).getCell(2).value != undefined && sh.getRow(i).getCell(4).value != null && sh.getRow(i).getCell(4).value != undefined  ) {

        const identifier = sh.getRow(i).getCell(1).value.text
        
        
         let sql = "SELECT * FROM tablename where col_name = ?"
         connection.query(sql,[identifier],function(err,result){
            if(result.length == 0){
            console.log(`insert into tablename (page_type, identifier, meta_title, meta_description, created_at, updated_at) values ("product", "${sh.getRow(i).getCell(1).value.text}", "${sh.getRow(i).getCell(2).value}", "${sh.getRow(i).getCell(4).value}", now(), now());`)
           sql  = `insert into tablename (page_type, identifier, meta_title, meta_description, created_at, updated_at) values ("product", "${sh.getRow(i).getCell(1).value.text}", "${sh.getRow(i).getCell(2).value}", "${sh.getRow(i).getCell(4).value}", now(), now());`
   
               connection.query(sql,function(err,res){
            console.log(res,"inserted value :",identifier)
             
       })   
      console.log("count_ins",++count_ins)
        }
         else{    
          for(let j=0;j<result.length;j++){
           console.log(`update tablename set meta_title = "${sh.getRow(i).getCell(2).value}" , meta_description = "${sh.getRow(i).getCell(4).value}" where identifier = "${sh.getRow(i).getCell(1).value.text}";`)
            sql = `update tablename set meta_title = "${sh.getRow(i).getCell(2).value}" , meta_description = "${sh.getRow(i).getCell(4).value}" where identifier = "${sh.getRow(i).getCell(1).value.text}";`
            connection.query(sql,function(err,res){
                console.log(res,"updated value :",identifier)
                 
           })   
           console.log("count_upt",++count_upt)
    }
}
             
         })
         
        
         }

    }
       
});
})
    }}
}
);