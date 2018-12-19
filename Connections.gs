/*
* CloudSimple 3.0
   by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
   v3.0   20181205 - Open Source Software
*/
function getActiveConnection(){
  var allConnections = JSON.parse(PropertiesService.getUserProperties().getProperty('CONNECTIONS'));
  for(var i = 0; i < allConnections.connections.length; i++){
    if(allConnections.connections[i].active){
      PropertiesService.getUserProperties().setProperty('AUTH', allConnections.connections[i].base64Auth);
      return allConnections.connections[i];
    }
  }

  //if we get here, there was no connection returned
  throw "No default connection set up";
}

function getConnections(){
  var allConnectionsString = PropertiesService.getUserProperties().getProperty('CONNECTIONS');
  if(allConnectionsString != null){
    var allConnections = JSON.parse(allConnectionsString);
    return allConnections;
  }
}

function saveConnections(connections){
  var allConnections = {connections:[]};
  for(var i = 0; i < connections.length; i++){
    if(connections[i].active == true){
      connections[i].base64Auth = Utilities.base64Encode(connections[i].sapUserName + ':' + connections[i].sapPassword);
      PropertiesService.getUserProperties().setProperty('AUTH', connections[i].base64Auth);
    }
    allConnections.connections[i] = connections[i];
  }

  PropertiesService.getUserProperties().setProperty('CONNECTIONS', JSON.stringify(allConnections));
}
