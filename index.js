const express = require("express")
const multer = require("multer")
const excelToJason = require("convert-excel-to-json")
const fs  = require("fs-extra")
const excelReader = require('xlsx')

const app = express()
const port = 3000


var  upload = multer({dest:"uploads/"})

app.post('/excel',upload.single('file'),(req,res)=>{
    var data = []
try {
    if (req.file?.filename == null || req.file?.filename == 'undefined') {
        res.status(400).json("No file...")
        
    } else {
        const filePath = 'uploads/'+req.file?.filename;
        const file   = excelReader.readFile(filePath)
        const sheetNames = file.SheetNames;
        
        for (let index = 0; index < sheetNames.length; index++) {
            const arr = excelReader.utils.sheet_to_json(file.Sheets[sheetNames[index]])
            arr.forEach((info)=>{
                data.push(info)
            })
            
        }
        res.status(200).json(data)
    }
    
} catch (error) {
   res.status(500).json({message: error.message}) 
}
})


app.listen(port, () => {
    console.log(`App listening on port ${port}`)
  }) 