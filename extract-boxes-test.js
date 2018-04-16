const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');

var workbook = XLSX.readFile('boxlist_ad.xlsx');
var worksheet = workbook.Sheets[workbook.SheetNames[0]];
var jsondata = XLSX.utils.sheet_to_json(worksheet); // box info stored as json-like object

var coordinates_array = [];
var promises_array = [];

var counter = 0;
var interval = setInterval(() => {
  if (counter >= jsondata.length) {
    clearInterval(interval);
    // putting this Promise.all outside of the setInterval method
    // leads to promises_array being empty
    Promise.all(promises_array).then((results) => {
      var failure_array = results.filter(x => x.status !== 'resolved');
      var success_array = results.filter(x => x.status === 'resolved');
      success_array.forEach((success) => {
        coordinates_array.push({
          address: success['address'],
          coordinates: success['v'],
          name: success['name'],
          local_name: success['local_name']
        });
      });
      fs.writeFileSync('results.json', JSON.stringify(coordinates_array));
      fs.writeFileSync('failure.json', JSON.stringify(failure_array));
    });

    return;
  }

  var box = jsondata[counter];
  if (!box) {return;}
  var encodedAddress = encodeURIComponent(box['Formatted 주소']);
  var url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodedAddress}`;
  promises_array.push(
    new Promise((resolve, reject) => {

      axios({
        method: 'get',
        url: url
      })
        .then(
        (value) => {
          if (value.data.status === 'OK') {
            resolve({
              status: 'resolved',
              address: value.data.results[0].formatted_address,
              v: value.data.results[0].geometry.location,
              name: box['CF Box Name'],
              local_name: box['Local Name']
            });
          } else {
            resolve({
              status: 'resolved but no data came back',
              name: box['CF Box Name'],
              local_name: box['Local Name']
            });
          }
      }, (reason) => {
        resolve({
          status: 'rejected',
          name: box['CF Box Name'],
          local_name: box['Local Name']
        });
      });

    })
  );
  console.log('Pushing a promise... ' + counter);
  counter++;
}, 1300);
