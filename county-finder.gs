function getGeocodingRegion() {
  return PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'us';
}

function zipToCounty(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  var addressColumn = 1;
  var addressRow;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  // Logger.log(geocoder.geocode("Baker, WV 26801, USA")); return;
  
  for(addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow){
    var address = cells.getCell(addressRow, addressColumn).getValue();
    var county = "";
    var city = "";
    var state = ""
    var cityState = "";
    
    location = geocoder.geocode(address);
    if (location.status == 'OK') {
      var address_components = location["results"][0]["address_components"];
      
      for(j=0; j < address_components.length; j++){
        if(address_components[j]["types"][0] == "administrative_area_level_2"){
          county = address_components[j]["short_name"];
        }
        
        if(address_components[j]["types"][0] == "locality"){
          city = address_components[j]["short_name"];
        }
        
        if(address_components[j]["types"][0] == "administrative_area_level_1"){
          state = address_components[j]["short_name"];
        }
      }
      cityState = city + ", " + state;
      location = geocoder.geocode(cityState);
      
      if( county ){
        cells.getCell(addressRow, addressColumn + 1).setValue(county);
      }else{
        var address_components = location["results"][0]["address_components"];
        for(j=0; j < address_components.length; j++){
          if(address_components[j]["types"][0] == "administrative_area_level_2"){
            county = address_components[j]["short_name"];
          }else if(address_components[j]["types"][0] == "locality"){
            county = address_components[j]["short_name"];
          }
        }
        cells.getCell(addressRow, addressColumn + 1).setValue(county);
      }
    }
  }
}

function zipToAdress(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  var addressColumn = 1; 
  var completeAddressCol = addressColumn + 1;
  var cityCol = completeAddressCol + 1;
  var countyCol = cityCol + 1;
  var addressRow;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  for(addressRow = 1; addressRow <= cells.getNumRows(); addressRow++){
    var address = cells.getCell(addressRow, addressColumn).getValue();
    location = geocoder.geocode(address);
    
    if (location.status == 'OK') {
      var formatted_address = location["results"][0]["formatted_address"];
      var address_components = location["results"][0]["address_components"];
      var county = "";
      var city = "";
      var state = "";
      
      for(j=0; j < address_components.length; j++){
        if(address_components[j]["types"][0] == "administrative_area_level_2"){
          county = address_components[j]["long_name"];
        }
        if(address_components[j]["types"][0] == "locality"){
          city = address_components[j]["long_name"];
        }
        if(address_components[j]["types"][0] == "administrative_area_level_1"){
          state = address_components[j]["long_name"];
        }
      }
      
      cells.getCell(addressRow, completeAddressCol).setValue(formatted_address);
      cells.getCell(addressRow, cityCol).setValue(city);
      if(county){
        cells.getCell(addressRow, countyCol).setValue(county);
      }else{
        location = geocoder.geocode(formatted_address);
        var address_components = location["results"][0]["address_components"];
        for(j=0; j < address_components.length; j++){
          if(address_components[j]["types"][0] == "administrative_area_level_2"){
            county = address_components[j]["short_name"];
          }else if(address_components[j]["types"][0] == "locality"){
            county = address_components[j]["short_name"];
          }
        }
        cells.getCell(addressRow, countyCol).setValue(county);
      }
    }
  }
}

function generateMenu() {
  var entries = [
    {
      name: "Zip to County",
      functionName: "zipToCounty"
    },
    {
      name: "Zip to Address",
      functionName: "zipToAdress"
    }
  ];
  
  return entries;
}

function updateMenu() {
  SpreadsheetApp.getActiveSpreadsheet().updateMenu('Google Geocode', generateMenu())
}


function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Google Geocode', generateMenu());
};
