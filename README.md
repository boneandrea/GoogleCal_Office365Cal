# GCal_365Cal

## Office365カレンダーの内容をGoogleカレンダーに転記します。
転記する期間は、実行時点からNヶ月間です。(default: 1)

### 実行方法

% ./run.sh [N|help]

### 実行前の設定
#### 365.yamlとcalendar.jsonを以下の内容で作成します。  
client_id, client_secretは https://console.developers.google.com/project?authuser=0  
access_tokenは https://developers.google.com/google-apps/calendar/auth
から調達されたし。

365.yaml:
```
office_id: *OfficeのID*  
office_pass: *Officeのpassword*  
```

calendar.json:

```
{  
"access_token" : "",  
"authorization_uri" : "https://accounts.google.com/o/oauth2/auth",  
"client_id" : "",  
"client_secret":""  
"expires_in":3600,  
"refresh_token":""  
"token_credential_uri":"https://accounts.google.com/o/oauth2/token",  
"issued_at":""  
}
```

#### ライブラリ設定  
% bundle install

#### 実行
% ./sync_task.rb [N]  


