function boldText(text_to_add) {
         let emailBody = document.getElementById("emailBody");
         let start_position = emailBody.selectionStart;
         let end_position = emailBody.selectionEnd;

         document.getElementById("emailBody").focus();

         emailBody.value = `${emailBody.value.substring(
                0,
                start_position
         )}${text_to_add}${emailBody.value.substring(
             end_position,
             emailBody.value.length
         )}`;

};
function italicText(text_to_add){
         let emailBody = document.getElementById("emailBody");
         let start_position = emailBody.selectionStart;
         let end_position = emailBody.selectionEnd;

         document.getElementById("emailBody").focus();
         emailBody.value = `${emailBody.value.substring(
                0,
                start_position
         )}${text_to_add}${emailBody.value.substring(
                end_position,
                emailBody.value.length
         )}`;
};
function underlineText(text_to_add){
          let emailBody = document.getElementById("emailBody");
          let start_position = emailBody.selectionStart;
          let end_position = emailBody.selectionEnd;

          document.getElementById("emailBody").focus();
          emailBody.value = `${emailBody.value.substring(
                0,
                start_position
          )}${text_to_add}${emailBody.value.substring(
                end_position,
                emailBody.value.length
          )}`;
};
function nextLine(text_to_add){
          let emailBody = document.getElementById("emailBody");
          let start_position = emailBody.selectionStart;
          let end_position = emailBody.selectionEnd;

          document.getElementById("emailBody").focus();
          emailBody.value = `${emailBody.value.substring(
                0,
                start_position
          )}${text_to_add}${emailBody.value.substring(
                end_position,
                emailBody.value.length
          )}`;
};
function validation(){
    if(document.emaildata.file.value.length < 35){
        document.getElementById("result").innerHTML="*CSV file - Please enter proper file directory.*";
        openErrorPopup();
        return false;
    }
    else if(document.emaildata.subject.value.length < 5){
        document.getElementById("result").innerHTML="*Subject - Please enter proper subject.*";
        openErrorPopup();
        return false;
    }
    else{
        openPopup();

    }
}
let popup = document.getElementById("popup");
function openPopup(){
//      if (document.emaildata.file.value.length > 35 && document.emaildata.subject.value.length > 5)
          document.getElementById("popup").classList.add("open-slide");
}
function closePopup(){
      document.getElementById("popup").classList.remove("open-slide");
}
let savepopup = document.getElementById("savepopup");
function openSavePopup(){
    if (document.emaildata.file.value.length > 35)
          document.getElementById("savepopup").classList.add("open-slide");
}
function closeSavePopup(){
    document.getElementById("savepopup").classList.remove("open-slide");
}
let errorpopup = document.getElementById("errorpopup");
function openErrorPopup(){
        document.getElementById("errorpopup").classList.add("open-slide");
}
function closeErrorPopup(){
    document.getElementById("errorpopup").classList.remove("open-slide")
}