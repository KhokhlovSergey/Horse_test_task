//Ссылка на табличку
//https://docs.google.com/spreadsheets/d/1aFeJCs4_f-C4JlQvrcSsWXEAwpRNZdMKx2nwZSF1tiU/edit#gid=0

function date_parse(date_income){
  //Функция возвращает дату в нужном формате
var date = new Date( date_income * 1000 );
var year = date.getFullYear();
var month = date.getMonth() + 1;
var day = date.getDate()
var hours = date.getHours();
var minutes = date.getMinutes(); 
var time = day + '.' + month + '.' + year + ' ' + hours + ':' + minutes;
return time;

}
//В скрипте присутствует триггер по времени. Каждые три часа автоматически запускается функция create_table, тем самым реализовывая функцию автообновления.

function create_table() {
  //получаем доступ к табличке по юрл, получаем доступ к листу
  var table = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1aFeJCs4_f-C4JlQvrcSsWXEAwpRNZdMKx2nwZSF1tiU/edit')
  var sheet = table.getSheets()[0];

  //Готовим таблицу
  sheet.getRange(1,1).setValue('id поста')
  sheet.getRange(1,2).setValue('дата')
  sheet.getRange(1,3).setValue('подпись к картинке')
  sheet.getRange(1,4).setValue('Картинка из поста')

  //Отправляем запрос к VK api, там предварительно было зарегистрированно приложение. Получаем джейсон, парсим его.
  //По запросу приходит десять последних постов (пожно больше) в сообществе. В качестве испытуемого был выбран Хабр-Карьера.
  var count = 10
  var filter = 'owner'
  var method = 'wall.get'
  var owner_id = '46638176'
  var response = UrlFetchApp.fetch('https://api.vk.com/method/' + method +'?owner_id=-' + owner_id + '&count=' + count + '&filter='+ filter +'&extended=0PARAMS&access_token=c54fc015c54fc015c54fc01571c532ae8bcc54fc54fc015a79d73d1e7eb769d4fe70bf9&v=5.131');
  var data = JSON.parse(response)['response']['items'];

  //В цикле перебирем все интересующие итемы, выводим их в табличку 
  for(var i =0;i != 10; i++){
      sheet.getRange(i + 2,1).setValue(i + 1);//выводим айдишник в таблице
      sheet.getRange(i + 2,2).setValue(date_parse(data[i].date)); //выводим дату публикации
      sheet.getRange(i + 2,3).setValue(data[i].text); //выводим подпись к посту
            
      if(data[i]['attachments'][0]['photo']){
        sheet.getRange(i + 2,4).setValue('=IMAGE("' + data[i]['attachments'][0]['photo']['sizes'][2].url + '";1)');// выводим изображение
      } else {
        sheet.getRange(i + 2,4).setValue('Нет изображения')
      }      
  }
}