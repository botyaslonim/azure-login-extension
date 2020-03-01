// инкапсуляция
(function(){	
	const key = "mKan82nJA01";	
	const state = "Om45nq1L";
	const authority = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
	const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
	const msGraphUrl = "https://graph.microsoft.com/v1.0/me";
	const clientId = "e2ed328f-6c55-4635-8bdb-d14880e7ee76";
	const redirectUri = "https://jomnjlcpondldnnnogpcpmekgckjdgkb.chromiumapp.org";
	const scope = "openid%20offline_access%20https%3A%2F%2Fgraph.microsoft.com%2Fmail.read";
	let appCode = null;	
	let accessToken = null;
	let refreshToken = null;
	let userName = null;
	let userAvatar = null;
	let gettingFreshToken = false;

	// запрашиваем информацию о пользователе с помощью токена, документация https://docs.microsoft.com/ru-ru/graph/auth-v2-user
	const getUserData = () => {
		if(!accessToken && !localStorage["azure_login_" + key + "_accessToken"]) return false; 
		fetch(msGraphUrl, 
		{
			method : "GET",
			headers : {
				"Authorization": "Bearer " + accessToken
			}		
		})
			.then(response => response.json())
			.then(response => {
				console.log(response)
				userName = response.displayName;
				fillData();				
			})
			// токен может протухнуть, пытаемся запросить новый
			.catch((error) => {				
				if (!gettingFreshToken) getFreshToken();
				gettingFreshToken = true;
			})
	}

	// вход в учетную запись
	const logIn = () => {
		chrome.identity.launchWebAuthFlow(
			{
				'url': authority + '?client_id=' + clientId + '&response_type=code&redirect_uri=' + redirectUri + '&response_mode=query&scope=' + scope + '&state=' + state, 
				'interactive': true
			},
			function(redirect_url) { 
				console.log("login redirect_url: " + redirect_url)
				let _st = redirect_url.slice(redirect_url.indexOf('state') + 6);
				if(_st === state) {
					appCode = redirect_url.slice(redirect_url.indexOf('code') + 5);
					appCode = appCode.slice(0, appCode.indexOf('&'));
					console.log(appCode);
					// сохраняем код доступа для проверки при следующем входе
					localStorage["azure_login_" + key + "_appCode"] = appCode;
					document.getElementById("logout").style.display = "block";
					document.getElementById("login").style.display = "none";

					// запрашиваем токен
					getToken(appCode);
				}				
			}
		);		
	};

	// получение токена с помощью кода доступа
	const getToken = (appCode) => {
		console.log("GET TOKEN")
		const body = JSON.stringify({
			"grant_type" : "authorization_code",
			"code" : appCode
		});	
		fetch(tokenUrl + "?client_id=" + clientId + "&scope=user.read%20mail.read&redirect_uri=" + redirectUri, 
			{
				method : "POST",
				body,
				headers : {
					"Content-Type": "application/json"
				}			
			})
			.then((response) => response.json())
			.then((response) => {
				console.log(response);
				accessToken = response.access_token;
				refreshToken = response.refresh_token;
				localStorage["azure_login_" + key + "_accessToken"] = accessToken;
				localStorage["azure_login_" + key + "_refreshToken"] = refreshToken;
			})
			.catch((error) => console.log(error))
	}	

	// получение свежего токена
	const getFreshToken = () => {
		const body = JSON.stringify({
			"grant_type" : "refresh_token",
			"code" : appCode
		});
		fetch(tokenUrl + "?client_id=" + clientId + "&scope=user.read%20mail.read&redirect_uri=" + redirectUri, 
			{
				method : "POST",
				body,
				headers : {
					"Content-Type": "application/json"
				}				
			})
			.then((response) => response.json())
			.then((response) => {
				console.log(response);
				accessToken = response.access_token;
				refreshToken = response.refresh_token;
				// новая пара токенов пришла, можем ещё раз спросить имя пользователя
				if (accessToken && refreshToken) getUserData();	 			
			})
			.catch((error) => console.log(error))
	}

	// заполнение данных о пользователе
	const fillData = () => {
		if (userName) {
			document.getElementById("user-name").innerHTML = userName;
			document.getElementById("user-name").style.display = "block";
		} else {
			document.getElementById("user-name").innerHTML = "";
			document.getElementById("user-name").style.display = "none";
		}
	};	

	// выход из учетной записи
	const logOut = () => {
		chrome.identity.launchWebAuthFlow(
			{		
				'url': 'https://login.windows.net/common/oauth2/v2.0/logout', 
				'interactive': true
			},
			function(redirect_url) { 
				console.log("logout redirect_url: " + redirect_url);	
				userName = null;
				accessToken = null;
				refreshToken = null;				
				localStorage["azure_login_" + key + "_appCode"] = "";
				localStorage["azure_login_" + key + "_accessToken"] = "";
				localStorage["azure_login_" + key + "_refreshToken"] = "";
				document.getElementById("logout").style.display = "none";
				document.getElementById("login").style.display = "block";
				fillData();
			}
		);
	}

	// при старте приложения запрашиваем данные о клиенте
	getUserData();

	// handlers
	document.getElementById("login").onclick = function() {	
		logIn();
	}
	document.getElementById("logout").onclick = function() {	
		logOut();
	}
})()