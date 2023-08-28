function verifyUser(){
 myInput = prompt("This feature is reserved for registered users. Software updates are available for downloading. Please enter your authorization code:", "*");
 if (myInput == "12354") {
  return true;
  }
 else {
  return false;
  }
}