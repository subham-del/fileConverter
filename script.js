let express=require('express');
let aspose = require('asposewordscloud')
const GIFEncoder = require('gifencoder');
const { createCanvas, loadImage } = require('canvas');
let app=express();
const path = require('path');
const pdf = require('pdf-poppler')
const p = require('pdf-page-counter');
var api = require('@asposecloud/aspose-html-cloud');
const fs = require('fs').promises;
const Fs=require('fs')
var AdmZip = require("adm-zip");
const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);
let multer=require('multer');
const {exec} = require('child_process');
const { stdout, stderr } = require('process');
const { clearScreenDown } = require('readline');
const notifier = require('node-notifier');

var fileDownload;
let storage=multer.diskStorage({
    destination:'./serverUploads',

    filename:(req,file,cb)=>
    {
        cb(null,file.originalname);
    }
})
let upload=multer({
    storage:storage
})

app.post("/pdfToWord",upload.single('file'),async(req,res)=>{
    const client = "08cff384-62b4-484d-99e6-ba250b724783"
    const api_key="2d6545d4194a8b303d4aec31f2ac8ac1"
    const filename=req.file.filename.split(".")
    if(filename[1]!='pdf'){
        notifier.notify({
            title: 'Alert',
            message: 'only pdf files are allowed'
          });
          res.redirect(req.get('referer'));
    }
    else{
        const outputFile = path.join(__dirname+ `/serverUploads/${filename[0]}.docx`);
        const wordsApi = new aspose.WordsApi(client,api_key);

        const doc = Fs.createReadStream(__dirname+ `/serverUploads/${req.file.filename}`);

        const request = new aspose.ConvertDocumentRequest({
        document: doc,
        format: "docx"
    });

    const convert = wordsApi.convertDocument(request)
    .then((result) => {
        
        
    
        Fs.writeFileSync(outputFile, result.body);
        
        fileDownload=outputFile;
        res.sendFile(__dirname+"/openfile.html");

    });
    }

  
})
    




app.post("/docxToPdf",upload.single('file'),async(req,res)=>{
    const ext = '.pdf'
    const filename=req.file.filename.split(".")
    console.log(filename[1])
    if(filename[1]!='docx' && filename[1]!='doc'){
        notifier.notify({
            title: 'Alert',
            message: 'only docx or doc files are allowed'
          });
          res.redirect(req.get('referer'));
    }
    else{
        const inputPath = path.join(__dirname+`/serverUploads/${req.file.filename}`);
        const outputPath = path.join(__dirname+ `/serverUploads/${filename[0]}.pdf`);
        const docxBuf = await fs.readFile(inputPath);
        let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
        await fs.writeFile(outputPath, pdfBuf);
        fileDownload=outputPath;
        res.sendFile(__dirname+"/openfile.html");
    }
    
})

app.post('/imageToPdf',upload.array('images',100),(req,res)=>{
   
    const filename=req.files[0].filename.split(".")
    
    if(filename[1]!='jpeg' && filename[1]!="png"){
        notifier.notify({
            title: 'Alert',
            message: 'only image files are allowed'
          });
          res.redirect(req.get('referer'));
    }
    else{
    const OutputPath =path.join(__dirname+`/serverUploads/${filename[0]}.pdf`);
    let list="";
    if(req.files){
        req.files.forEach(file =>{
           let paths = path.join(__dirname+`/${file.path}`)
            list+=`"${paths}"`
            list+=" "
        })
    }
   
    exec(`magick convert ${list} "${OutputPath}"`,(err,stdout,stderr)=>{
       
        if(err){
            throw err;
        }
        else{
            fileDownload=OutputPath;
            res.sendFile(__dirname+"/openfile.html");
           
        }
    
    })
}
})


app.get('/download',(req,res)=>{
    res.download(fileDownload)
})



app.post('/pngTojpeg',upload.single('image'),(req,res)=>{
    const filename=req.file.filename.split(".")
    if(filename[1]!="png"){
        notifier.notify({
            title: 'Alert',
            message: 'only png files are allowed'
          });
          res.redirect(req.get('referer'));
         
    }
    else{
    const OutputPath =path.join(__dirname+`/serverUploads/${filename[0]}.jpeg`);
    const list=path.join(__dirname+`/serverUploads/${req.file.filename}`);
    exec(`magick convert ${list} ${OutputPath}`,(err,stdout,stderr)=>{
       
        if(err){
            throw err;
        }
        else{
            fileDownload=OutputPath;
            res.sendFile(__dirname+"/openfile.html");
           
        }
    
    
    })
}
})


app.post('/htmlToPdf',upload.single('file'),(req,res)=>{

    const filename=req.file.filename.split(".")
    if(filename[1]!="html"){
        notifier.notify({
            title: 'Alert',
            message: 'only HTML files are allowed'
          });
          res.redirect(req.get('   referer'));
         
    }
    else{
    var conf = {
        "basePath":"https://api.aspose.cloud/v4.0",
        "authPath":"https://api.aspose.cloud/connect/token",
        "apiKey":"2d6545d4194a8b303d4aec31f2ac8ac1",
        "appSID":"08cff384-62b4-484d-99e6-ba250b724783",
        "defaultUserAgent":"NodeJsWebkit"
    };
    
   
    
    // Create Conversion Api object
    var conversionApi = new api.ConversionApi(conf);
    
    var src = __dirname+`/serverUploads/${req.file.filename}`; // {String} Source document.
    var dst = __dirname+`/serverUploads/${filename[0]}.pdf`  // {String} Result document.
    var opts = null;
    
    var callback = function(error, data, response) {
      if (error) {
        console.error(error);
      } else {
       
        fileDownload = __dirname+`/serverUploads/${filename[0]}.pdf`
        res.sendFile(__dirname+"/openfile.html")
      }
    };
    
    conversionApi.convertLocalToLocal(src, dst, opts, callback);
}
   
})





app.post('/Mp4ToMp3',upload.single('file'),(req,res)=>{
    
    const filename=req.file.filename.split(".")
   
    if(filename[1]!="mp4"){
        notifier.notify({
            title: 'Alert',
            message: 'only mp4 files are allowed'
          });
          res.redirect(req.get('referer'));
         
    }
    else{
   
    const inputPath = path.join(__dirname+`/serverUploads/${req.file.filename}`);
    const OutputPath =path.join(__dirname+`/serverUploads/${filename[0]}.mp3`);
  
    exec(`ffmpeg -i "${inputPath}" "${OutputPath}"`, (error, stdout, stderr) => {
        if (error) {
            console.log(`error: ${error.message}`);
            return;
        }
        else{
            console.log("file is converted")
            fileDownload=OutputPath;
            res.sendFile(__dirname+"/openfile.html");
       
    }
    })
}
})





app.post('/pdfToJpg',upload.single('file'),(req,res)=>{
    const filename=req.file.filename.split(".")
    if(filename[1]!='pdf'){
        notifier.notify({
            title: 'Alert',
            message: 'only pdf files are allowed'
          });
          res.redirect(req.get('referer'));
    }
    else{
    const zip = new AdmZip();
    let dataBuffer = Fs.readFileSync(path.join(__dirname+`/serverUploads/${req.file.filename}`));
    p(dataBuffer).then(function(data) {
        for(let i=1;i<=data.numpages;i++){
    const list=path.join(__dirname+`/serverUploads/${req.file.filename}`);
    let option = {
        format : 'jpeg',      
        out_dir : './serverUploads',
        out_prefix : path.basename(list, path.extname(list)),
        page : `${i}`
    }
// option.out_dir value is the path where the image will be saved

    pdf.convert(list, option)
    .then(() => {
       
    })
    .catch(err => {
        console.log('an error has occurred in the pdf converter ' + err)
    })
}
    
    })
   
     setTimeout(function () {

        Fs.readdir('./serverUploads', (err, files) => {

                const jpgFiles = files.filter((e1)=> path.extname(e1) === ".jpg")
               
                console.log(jpgFiles)
                jpgFiles.forEach(file=>{
                    zip.addLocalFile(`./serverUploads/${file}`)
                })
                Fs.writeFileSync('./outputFiles/output.zip',zip.toBuffer())

                fileDownload = './outputFiles/output.zip'
                res.sendFile(__dirname+"/openfile.html");
          });
         
      
        
    },2500);  
}
})





app.post("/gifToimages",upload.single('file'),(req,res)=>{
    const filename=req.file.filename.split(".")
    if(filename[1]!='gif'){
        notifier.notify({
            title: 'Alert',
            message: 'only gif files are allowed'
          });
          res.redirect(req.get('referer'));
    }
    else{
    const inputPath = path.join(__dirname+`/serverUploads/${req.file.filename}`)
        const outputPath = path.join(__dirname+ `/serverUploads/frame%d.jpg`);
        exec(`ffmpeg -i ${inputPath} -vsync 0 ${outputPath}`, (error, stdout, stderr) => {
            if (error) {
                console.log(`error: ${error.message}`);
                return;
            }
            else{
                res.sendFile(__dirname+"/thank_you.html");
           
        }
        })
    
    }
})




app.get('/',(req,res)=>{
    res.sendFile(__dirname+"/for_screen.html")
})

app.get('/:id',(req,res)=>{
    res.sendFile(__dirname+`/${req.params.id}`)
})

                                                                                                                                                                                                                                                                                                                                                                                                                 
  
app.get('/openfile.html',(req,res)=>{
    res.sendFile(__dirname+"/openfile.html")
})

app.get('/files/:id',(req,res)=>{
    console.log(req.params.id)
    res.sendFile(__dirname+`/serverUploads/${req.params.id}`)
})
app.get('/style.css',(req,res)=>{
    res.sendFile(__dirname+"/style.css")
})
app.get('/docx_to_pdf.html',(req,res)=>{
    res.sendFile(__dirname+"/docx_to_pdf.html")
})
app.get('/image_to_pdf.html',(req,res)=>{
    res.sendFile(__dirname+"/image_to_pdf.html")
})

app.get('/png_to_jpeg.html',(req,res)=>{
    res.sendFile(__dirname+"/png_to_jpeg.html")
})



app.listen(3000,()=>{console.log(`server is running at port 3000`);})