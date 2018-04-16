const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');

var workbook = XLSX.readFile('boxlist_ad.xlsx');
var worksheet = workbook.Sheets[workbook.SheetNames[0]];
var jsondata = XLSX.utils.sheet_to_json(worksheet); // box info stored as json-like object

var coordinates_array = [];
var promises_array = [];

// needed to deal with cases where certain promises get rejected
function reflect(promise) {
  return promise.then((value) => {
    if (value.data.status === 'OK') {
      return {
        address: value.data.results[0].formatted_address,
        v: value.data.results[0].geometry.location,
        status: 'resolved'};
    } else {
      return {status: 'resolved but no data came back'}
    }
  }, (error) => {
    return {e: error, status: 'rejected'};
  });
}

var counter = 0;
var interval = setInterval(() => {
  if (counter >= jsondata.length) {
    clearInterval(interval);
    // putting this Promise.all outside of the setInterval method
    // leads to promises_array being empty
    Promise.all(promises_array.map(reflect)).then((results) => {
      var success_array = results.filter(x => x.status === 'resolved');
      success_array.forEach((success) => {
        coordinates_array.push({
          address: success['address'],
          coordinates: success['v']
        });
      });
      fs.writeFileSync('results.json', JSON.stringify(coordinates_array));
    });

    return;
  }
  var box = jsondata[counter];
  if (!box) {return;}
  var encodedAddress = encodeURIComponent(box['Formatted 주소']);
  var url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodedAddress}`;
  promises_array.push(axios({
    method: 'get',
    url: url
  }));
  console.log('Pushing a promise... ' + counter);
  counter++;
}, 1300);
