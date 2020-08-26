//requirements
const express = require('express');
const PORT = 3000;
const app = express();
const Shell = require('node-powershell');

//initialize a shell instance
const ps = new Shell({
    executionPolicy: 'Bypass',
    noProfile: true
});


//PS Request to site 
app.use('/', (req, res) => {
    // ps.addCommand(`Get-Process | ? { $_.name -like '*chrome*' }`);
    ps.addCommand("Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias *wi-fi* | ConvertTo-Json");
    ps.invoke()
        .then(response => {
        	// console.log(typeof response)
        	let newString = JSON.parse(response)
        	// console.log(newString)
            res.json(newString)
        })
        .catch(err => {
            res.json(err)
        });
});

//Listener
app.listen(PORT, ()  => {
	console.log('Lisenting on localhost:' + PORT)
});