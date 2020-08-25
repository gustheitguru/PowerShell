//requirements
const express = require('express');
const PORT = 3000;
const app = express();
const Shell = require('node-powershell');

//initialize a shell instance
const ps = new Shell({
    executionPolicy: 'Bypass',
    noProfile: true


//PS Request to site 
app.use('/', (req, res) => {
    ps.addCommand(`Get-Process | ? { $_.name -like '*chrome*' }`);
    ps.invoke()
        .then(response => {
        	console.log(response)
            res.json(response)
        })
        .catch(err => {
            res.json(err)
        });
});

//Listener
app.listen(PORT, ()  => {
	console.log('Lisenting on localhost:' + PORT)
});