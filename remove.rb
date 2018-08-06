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

$CONF_FILE = "365.yaml"
$MYDEBUG = 1

class MyGCal
  @@CREDENTIAL_STORE_FILE = "calendar.json"
  @@client = nil
  @@entries = []
  @@service = nil
  @@config = nil

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
    @@config = YAML.load_file(File.dirname(__FILE__) + "/" + $CONF_FILE)

    $OFFICE_ID = @@config["office_id"]
    $OFFICE_PASS = @@config["office_pass"]
  end

  def auth_google
    file_storage = Google::APIClient::FileStorage.new(File.dirname(__FILE__) + "/" + @@CREDENTIAL_STORE_FILE)

    if file_storage.authorization.nil?
      client_secrets = Google::APIClient::ClientSecrets.load
      @@client.authorization = client_secrets.to_authorization
      @@client.authorization.scope = 'https://www.googleapis.com/auth/calendar'
    else
      @@client.authorization = file_storage.authorization
    end
  end

  ## 指定期間中のGoogleカレンダーイベント取得　＋　指定ワードで消す

  def get_google_tasks(timeMin, timeMax, delWord)
    @@service = @@client.discovered_api('calendar', 'v3')

    page_token = nil

    result = @@client.execute(:api_method => @@service.events.list,
                              :parameters => {'calendarId' => 'primary',
                                              'timeMin' => timeMin.rfc3339,
                                              'timeMax' => timeMax.rfc3339,
                                              :maxResults => 2500
                                             })

    # if is_error_response(result) then
    #   pp result.body
    #   error_response=JSON.parse(result.body["error"])

    #   error_mail(error_response["code"])

    #   return false
    # end
    @@entries.concat(result.data.items)

    catch(:exit) do
      while true

        throw :exit

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

    @@entries.each do |e|
      if e.summary == delWord
      then
        p "delete one"
        delete(e)
      end
    end
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
    options = { :address => @@config["smtp_server"],
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
	    body "Oh,no"
    end

    mail.deliver!
  end

  def myp(s)
    puts s if $MYDEBUG == 1
  end

  def delete(e)
    result = @@client.execute(:api_method => @@service.events.delete,
                              :parameters => {'calendarId' => 'primary', 'eventId' => e.id})
  end
end

term_month = if ARGV[1] == "help" then
               help = true
               nil
             else
               help = false
               # default: 今日から1ヶ月だけ
               (ARGV[1] || 1).to_i
             end

x = MyGCal.new

if help then
  x.help
  exit 0
end

timeMax = Date.today
timeMin = Date.today - 7

p timeMin
p timeMax

x.get_google_tasks(timeMin, timeMax, ARGV[0])

exit 0
