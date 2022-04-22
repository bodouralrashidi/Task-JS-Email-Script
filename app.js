// Your code goes here

const excelToJson = require('convert-excel-to-json');
 
const result = excelToJson({
    sourceFile: 'names.xlsx'
});


let resultjson = result["Sheet1"]
resultjson.forEach(element => {

sendEmail(element.A,element.B,element.C)
  
});

//sendEmail("Bodour", "bodouralrashidi@gmail.com", "blahblah")
 // using Twilio SendGrid's v3 Node.js Library
// https://github.com/sendgrid/sendgrid-nodejs
function sendEmail(name , email, certificate){
  const sgMail = require('@sendgrid/mail')
  sgMail.setApiKey(process.env.SENDGRID_API_KEY)
const msg = {
  to: email, // Change to your recipient
  from: email, // Change to your verified sender
  subject: 'Certificate',
  text: 'dd',
  html: `<div style="width:800px; height:600px; padding:20px; text-align:center; border: 10px solid #000000">
  <div style="width:750px; height:550px; padding:20px; text-align:center; border: 5px solid #000000">
         <span style="font-size:50px; font-weight:bold">Special Awards</span>
         <br><br>
         <span style="font-size:25px"><i>presented to</i></span> <br>
         <span style="font-size:30px"><b>${name}</b></span><br/><br/>
         <span style="font-size:30px">${certificate}</span> <br/><br/>
         <span style="font-size:25px"><i>${getdate()}</i></span><br>
  </div>
  </div>`,
}
sgMail
  .send(msg)
  .then(() => {
    console.log('Email sent')
  })
  .catch((error) => {
    console.error(error)
  })

}

function getdate()
{
    var date = new Date();

var day = date.getDate();
var month = date.getMonth() + 1;
var year = date.getFullYear();

if (month < 10) month = "0" + month;
if (day < 10) day = "0" + day;

return year + "/" + month + "/" + day;       

  }