jQuery(document).ready(function() {
    // valid condition for show textarea input
  if($('form input[value="calling"]').is(':checked')) {
    $('textarea').css('display', 'block');
  }else{
    $('textarea').css('display', 'none');
  }

  // // valid condition for show datetime input
  // if($('form input[value="done"]').is(':checked')) {
  //   $('.date').css('display', 'none');
  // }else{
  //   $('.date').css('display', 'none');
  // }
});


function text_area_check(input){
  // valid condition for show textarea input
  if (input.attr('value')=='calling' ){
    $('textarea').css('display', 'block');
  }else{
    $('textarea').css('display', 'none');
  }
}
