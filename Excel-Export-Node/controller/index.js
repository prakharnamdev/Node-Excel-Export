const MongoClient = require('mongodb').MongoClient;
const url = "mongodb://localhost:27017/excel-import";
const excel = require('exceljs');

const homePage=(req,res)=>{
    res.render('index');
}

const exportExcel=(req,res)=>{

    // Create a connection to the MongoDB database
    MongoClient.connect(url, { useNewUrlParser: true }, function(err, db) {
      if (err) throw err;
      
      let dbo = db.db("excel-import"); //db name
      
      dbo.collection("users").find({}).toArray(function(err, result) { //student= table name
        if (err) throw err;
        console.log(result);	
        let workbook = new excel.Workbook(); //creating workbook
        let worksheet = workbook.addWorksheet('Users'); //creating worksheet

        //  WorkSheet Header
        worksheet.columns = [
            { header: 'Id', key: '_id', width: 10 },
            { header: 'Name', key: 'name', width: 30 },
            { header: 'Email', key: 'email', width: 30},
            { header: 'Phone', key: 'phone', width: 10},
            { header: 'Address', key: 'address', width: 20},
            { header: 'Age', key: 'age', width: 20, outlineLevel: 1}
        ];
        
        // Add Array Rows
        worksheet.addRows(result);
        
        // Write to File
         workbook.xlsx.writeFile("user.xlsx")
            .then(function() {
                console.log("file saved!");
              
            });
            db.close();
        
        
      });
    });
    res.send('Excel Export Successfully')
}

module.exports={exportExcel,homePage}