/*
========================================
Route.gs
負責頁面路由
========================================
*/

function doGet(e){

  let page = '';

  if(
    e &&
    e.parameter &&
    e.parameter.page
  ){
    page =
      e.parameter.page;
  }

  const adminPageKey =
    ADMIN_PAGE_PREFIX +
    ADMIN_ACCOUNT +
    ADMIN_PASSWORD;


  /*
  ========================================
  會議簽到
  ========================================
  */

  if(page === 'signmeeting'){
    return showSignMeetingPage(e);
  }


  /*
  ========================================
  QR產生
  ========================================
  */

  if(page === 'signqr'){
    return showSignQRGeneratorPage();
  }


  /*
  ========================================
  醫院系統登入
  ========================================
  */

  if(page === 'hospitalsignin'){
    return showHospitalSignInPage();
  }


  /*
  ========================================
  後台首頁
  ========================================
  */

  if(page === adminPageKey){
    return showAdminIndexPage();
  }


  /*
  ========================================
  會議管理後台
  ========================================
  */

  if(
    page ===
    adminPageKey + 'meeting'
  ){
    return showAdminMeetingPage();
  }


  /*
  ========================================
  預設首頁
  ========================================
  */

  return showFrontIndex();

}


// 從 html 調用其他 html/script
function include(filename){

  return HtmlService
    .createHtmlOutputFromFile(
      filename
    )
    .getContent();

}