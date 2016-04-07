#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'rubygems'
require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/file_storage'

require 'open-uri'
require "pp"

require 'yaml'
require 'mail'

require "date"
require "time"


$CONF_FILE="365.yaml"
$MYDEBUG=0

class MyGCal

  @@CREDENTIAL_STORE_FILE = "calendar.json"
  @@client = nil
  @@entries=[]
  @@service=nil
  @@config=nil
  
  def initialize
    @@client = Google::APIClient.new(
      :application_name => 'Ruby Calendar sample',
      :application_version => '1.0.0')

    begin
      setup
      auth_google
    rescue Exception => e
      p e
      exit 1
    end
    
    
  end

  ## Google/365 の認証情報設定
  ## ファイルから読む, 失敗 -> 例外

  def setup

    @@config = YAML.load_file(File.dirname(__FILE__)+"/"+$CONF_FILE)
    
    $OFFICE_ID=@@config["office_id"]
    $OFFICE_PASS=@@config["office_pass"]

  end


  def auth_google
    
    file_storage = Google::APIClient::FileStorage.new(File.dirname(__FILE__)+"/"+@@CREDENTIAL_STORE_FILE)
    
    if file_storage.authorization.nil?
      client_secrets = Google::APIClient::ClientSecrets.load
      @@client.authorization = client_secrets.to_authorization
      @@client.authorization.scope = 'https://www.googleapis.com/auth/calendar'
    else
      @@client.authorization = file_storage.authorization
    end

  end
  
  
  ## 指定期間中のGoogleカレンダーイベント取得

  def get_google_tasks(timeMin, timeMax)
    
    @@service = @@client.discovered_api('calendar', 'v3')
    
    page_token = nil
    
    result = @@client.execute(:api_method => @@service.events.list,
                              :parameters => {'calendarId' => 'primary',
                                              'timeMin'=>timeMin.rfc3339,
                                              'timeMax'=>timeMax.rfc3339
                                             })

    begin
      if is_error_response(result) then
        pp result.body
        error_response=JSON.parse(result.body["error"])
        error_mail(error_response["code"])
        return false
      end
    rescue => e

      return false

    end

    catch(:exit) do
      while true
        
        @@entries.concat(result.data.items)

        if !(page_token = result.data.next_page_token)
          throw :exit
        end
        result = @@client.execute(:api_method => @@service.events.list,
                                  :parameters => {'pageToken' => page_token})

        if is_error_response(result) then

          error_mail(result.body["error"]["code"] + result.body["error"]["code"])

          return false
        end
      end
    end

    print "#{@@entries.size} entries\n"

    true

  end
  
  def is_error_response(r)
    JSON.parse(r.body).has_key?("error")

    # TODO: 予定に2種類ある。終日と時刻ありと。
#    "start"=>{"dateTime"=>"2015-05-22T17:00:00+09:00"},
#    "end"=>{"dateTime"=>"2015-05-22T18:00:00+09:00"},

#    "start"=>{"date"=>"2015-05-22"},
#    "end"=>{"date"=>"2015-05-22"},

    
  end
  
  def error_mail(error_msg)
      
      options = { :address              => @@config["smtp_server"],
		          :port                 => @@config["smtp_port"],
		          :authentication       => @@config["smtp_auth"],
                  :enable_starttls_auto => @@config["enable_starttls_auto"],
		          :ssl => @@config["ssl"]
                }


      Mail.defaults do
		delivery_method :smtp, options
      end

      mail = Mail.new do
	    from     "365@peixe.biz"
	    to       "banchou@peixe.biz"
	    subject  error_msg
	    body    "Oh,no"
      end

      mail.deliver!

  end
  
  ## 365のイベントで、Googleカレンダーに含まれていなければinsert処理を呼ぶ

  def sync(office_events)

    office_events.each do |e|
      if check_and_insert(e) then
        p e
      else
        p "NO"
      end
    end
    
  end
        

  def myp(s)
    if $MYDEBUG == 1 then
      puts s
    end
  end

  def is_contained(task_name, start_time)
    @@entries.each do |e|

      mydate= (e.start["date"] || e.start["dateTime"]).to_s
      

      if e.summary == task_name && compare_date(mydate, start_time) then
        p "HIT! DUP!"
        return true
      else
        myp "#{e.summary} == #{task_name} && #{mydate}, #{start_time}"
      end

    end
    false
  end

  ## insert に失敗したらfalse, それ以外はtrue

  def check_and_insert(arg)

    task_name=arg[:subject]
    start_time=arg[:start]
    end_time=arg[:end]

    if is_contained(task_name, start_time) then
      return true
    else

      puts "INSERT"
      event = {
        'summary' => task_name,
        'location' => 'Somewhere',
      }

      if(arg[:isAllDay]) then
        event["start"]={
          'date' => start_time.to_s.sub(/T.*/,""),
          "time_zone"=>"Asia/Tokyo"
        }
        event["end"]={
          'date' => end_time.to_s.sub(/T.*/,""),
          "time_zone"=>"Asia/Tokyo"
        }
      else
        event["start"]={
          'dateTime' => start_time.to_s,
          "time_zone"=>"Asia/Tokyo"
        }
        event["end"]={
          'dateTime' => end_time.to_s,
          "time_zone"=>"Asia/Tokyo"
        }
      end

      pp "OH"
      pp event
      
      begin
        result = @@client.execute(:api_method => @@service.events.insert,
                                  :parameters => {'calendarId' => 'primary'},
                                  :body => JSON.dump(event),
                                  :headers => {'Content-Type' => 'application/json'})

        json=JSON.parse(result.body)
        p json
        return json["summary"] == task_name
        
      rescue Exception => e

        p e
        
      end
    end
    false
  end

  ##日付の比較関数
  
  def compare_date(d1,d2)

    Time.parse(d1) == Time.parse(d2)
    
  end

  ## 指定期間内の365のイベントリスト取得

  def get_office_tasks(timeMin, timeMax)

    p "365 CAL: " + timeMin.to_s + " -> "+timeMax.to_s
    uri="https://outlook.office365.com/api/v1.0/me/calendarview?startDatetime=#{timeMin.to_s}&endDateTime=#{timeMax.to_s}";

    certs =  [$OFFICE_ID, $OFFICE_PASS]
    json = JSON.parse(open(uri, {:http_basic_authentication => certs}).read)
    json["value"].map do |t|

      {
        :subject=> t["Subject"],
        :start=> Time.parse(t["Start"]).localtime.iso8601.to_s,
        :end=> Time.parse(t["End"]).localtime.iso8601.to_s,
        :isAllDay => t["IsAllDay"]
      }
    end

  end


  def help
    msg=<<-"AAA"
[Officeカレンダー -> Googleカレンダー同期の仕様]

usage: ./add_task.rb [N]

今日からNヶ月間において、
Officeカレンダーにあって,Googleカレンダーにない予定をGoogleカレンダーに登録する。
N は整数。引数で与える。デフォルトは1。
AAA

    STDERR.puts(msg)
      
    exit 0

  end

end
             

if ARGV[0] == "help" then
  help=true
  term_month = nil
else
  help=false
# default: 今日から1ヶ月だけ
  term_month = (ARGV[0] || 1).to_i
end


p Time.now
x=MyGCal.new

if help then
  x.help
  exit 0
end



timeMax=Date.today >> term_month
timeMin=Date.today - 1 # 全日予定をとるため -1 する

## Officeのタスク一覧を取ってきて
office_tasks=x.get_office_tasks(timeMin, timeMax)

## Googleのタスク一覧を取ってきて
if x.get_google_tasks(timeMin, timeMax) then
  ## 比較してGoogleにINSERT
  x.sync(office_tasks)

else
  p "google failed"
end

exit 0
