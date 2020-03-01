1. Создаём расширение Chrome Extension. 
Для этого в отдельной папке создаётся файл manifest_version.json с перечислением настроек и указанием html-файла с вёрсткой попапа.
Там же создаём html-файл, css- и js-файлы, картинку с иконкой расширения.
В браузере Chrome в панели управления расширениями включаем режим разработчика, подгружаем содержимое папки. Расширение устанавливается автоматически.
Расширение работает как обычная html-страница. Консоль браузера для этой страницы доступна, если кликнуть правой кнопкой мыши по области попапа. Логирование ошибок осуществляется в панели управления расширениями chrome://extensions/

2. Аккаунт Microsoft Azure
Создаём аккаунт в Azure, регистрируем новое прложение Web.
В панели управления смотрим client_id приложения.

3. Согласно документации Chrome extensions https://developer.chrome.com/apps/app_identity#non:
 - регистрируем провайдера не-гугловой аутентификации (в нашем случае Azure), в панели управление Azure Active directory записываем redirect_url следующим образом: https://jomnjlcpondldnnnogpcpmekgckjdgkb.chromiumapp.org/
 - добавляем разрешение на адрес Azure в manifest.json/permissions : https://login.microsoftonline.com/common/oauth2/v2.0/authorize
 - получаем токен авторизации, в main.js нашего приложения выполняем следующее:
chrome.identity.launchWebAuthFlow(
  {'url': (auth-url), 'interactive': true},
  function(redirect_url) { /* Extract token from redirect_url */ });

URL авторизации задаём согласно правилам Microsoft: https://docs.microsoft.com/ru-ru/azure/active-directory/develop/v2-oauth2-auth-code-flow

После прохождения регистрации сработает коллбэк с указанием страницы приложения и кода доступа в параметрах, а также строки state для проверки: 
https://jomnjlcpondldnnnogpcpmekgckjdgkb.chromiumapp.org/?code=....&state=12345
Проверяем соответствие state и записываем код.

Далее запрашиваем токен, при успешном получении можно запрашивать данные о пользователе.

4. Обновление токенов и рестарт приложения
Токен быстро протухает. В случае неуспешного запроса информации по токену пробуем его обновить. При рестарте приложения не запрашиваем данные пользователя, если токена нет


