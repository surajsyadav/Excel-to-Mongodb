var express    = require('express');
var mongoose   = require('mongoose');
var bodyParser = require('body-parser');
var path       = require('path');
var XLSX       = require('xlsx');
var async =require('async')
var multer     = require('multer');
//multer
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/upload')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
  });
  
  var upload = multer({ storage: storage });

//connect to db
mongoose.connect('mongodb://localhost:27017/Demoexcel',{useNewUrlParser:true})
.then(()=>{console.log('connected to db')})
.catch((error)=>{console.log('error',error)});

//init app
var app = express();

//set the template engine
app.set('view engine','ejs');

//fetch data from the request
app.use(bodyParser.urlencoded({extended:false}));

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));

//collection schema
var excelSchema = new mongoose.Schema({
   Name:String,
   Email:{
    type:String,
    required:true,
    unique: true
   },
   Mobile:String,
   DOB:String,
   workExperience:String,
   resumeTitle:String,
   currentLocation:String,
   postalAddress:String,
   currentEmployer:String,
   currentDesignation:String
});

var excelModel = mongoose.model('excelData',excelSchema);


app.get('/',(req,res)=>{
res.render('home');
});

app.post('/',upload.single('excel'),(req,res)=>{
  var workbook =  XLSX.readFile(req.file.path,{dateNF:"dd mmm yyyy"});
  var sheet_namelist = workbook.SheetNames;

    var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[0]],{raw:false});
    
    async.eachSeries(xlData,  function (element,callback) {

        callback(null,
          excelModel.insertMany(element,{ordered:false},(err,data)=>{
          if(err){
              console.log(err);
          }else{
              console.log(data);
          }
      }));
       
       
      }) 
      
     
    res.render('success');
});

//assign port
var port = process.env.PORT || 3000;
app.listen(port,()=>console.log('server run at '+port));