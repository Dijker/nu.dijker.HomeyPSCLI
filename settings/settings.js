//let Homey;

function onHomeyReady( homeyReady ){
    Homey = homeyReady;
    Homey.ready();
}

function SetCookie(c_name,value,expiredays) {
  var exdate=new Date()
  exdate.setDate(exdate.getDate()+expiredays)
  document.cookie=c_name+ "=" +escape(value)+
  ((expiredays==null) ? "" : ";expires="+exdate.toGMTString())+"; path=/"
  location.reload()
}
