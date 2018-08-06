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

$CONF_FILE = "365.yaml"

class MyGCal
  @@CREDENTIAL_STORE_FILE = "calendar.json"
  @@client = nil
  @@google_entries = []
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

  ## 指定期間中のGoogleカレンダーイベント取得

  def get_google_tasks(timeMin, timeMax)
    @@service = @@client.discovered_api('calendar', 'v3')

    page_token = nil

    # TODO: これに書き換えたい
    # begin
    #   result = @@client.list_events('primary')
    #   result.items.each do |e|
    #     print e.summary + "\n"
    #   end
    #   if result.next_page_token != page_token
    #     page_token = result.next_page_token
    #   else
    #     page_token = nil
    #   end
    # end while !page_token.nil?

    # exit

    begin
      result = @@client.execute(:api_method => @@service.events.list,
                                :parameters => {'calendarId' => 'primary',
                                                'timeMin' => timeMin.rfc3339,
                                                'timeMax' => timeMax.rfc3339,
                                                :maxResults => 2500
                                               })
    rescue => e
      p e
      exit
    end

    @@google_entries.concat(result.data.items)

    return true

    # ############### これ以降はいっぱい予定がある場合

    catch(:exit) do
      while true

        if !(page_token = result.data.next_page_token)
          throw :exit
        end

        begin
          result = @@client.execute(:api_method => @@service.events.list,
                                    :parameters => {'pageToken' => page_token,
                                                :maxResults => 2500
                                                   })

          @@google_entries.concat(result.data.items)
          if is_error_response(result) then

            error_mail(result.body["error"]["code"] + result.body["error"]["code"])

            return false
          end
        rescue => e
          # 最後まで行ったので握りつぶす
          break
        end
      end
    end

    print "#{@@google_entries.size} entries\n"
    @@google_entries.each do |e|
#      pp e
    end

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

  ## 365のイベントで、Googleカレンダーに含まれていなければinsert処理を呼ぶ

  def sync(office_tasks)
    office_tasks.each do |e|
      p e if check_and_insert(e)
    end
  end

  def is_contained(arg)
    office_task_name = arg[:subject]
    start_time = arg[:start]

    # s1="2016-04-08T20:00:01+09:00"
    # s2="2016-04-08T11:00:00+00:00"

    # if(DateTime.parse(s1) == DateTime.parse(s2))
    #   p "SAME"
    # else
    #   p "S!!!AME"
    # end

    @@google_entries.each do |g|
      next unless g.start

      if g.start["dateTime"] then
        mydate = Time.parse(g.start["dateTime"].to_s)
        start_time = Time.parse(g.start["dateTime"].to_s)
      else
        mydate = Time.parse(g.start["date"])
        start_time = Time.parse(g.start["date"])
      end

      # p "TYT #{g.summary}/#{arg[:subject]}"
      # p mydate
      # p start_time

      if g.summary == office_task_name && (mydate == start_time) then
        p "HIT! DUP!"
        return true
      end
    end

    false
  end

  ## insert に失敗したらfalse, それ以外はtrue

  def check_and_insert(arg)
    start_time = arg[:start]
    end_time = arg[:end]

    if is_contained(arg) then
      return true
    else

      puts "INSERT"
      event = {
        'summary' => arg[:subject],
        'location' => 'Somewhere',
      }

      if arg[:isAllDay] then
        event["start"] = {
          'date' => start_time.to_s.sub(/T.*/, ""),
          "time_zone" => "Asia/Tokyo"
        }
        event["end"] = {
          'date' => end_time.to_s.sub(/T.*/, ""),
          "time_zone" => "Asia/Tokyo"
        }
      else
        event["start"] = {
          'dateTime' => start_time.to_s,
          "time_zone" => "Asia/Tokyo"
        }
        event["end"] = {
          'dateTime' => end_time.to_s,
          "time_zone" => "Asia/Tokyo"
        }
      end

      pp event

      begin
        result = @@client.execute(:api_method => @@service.events.insert,
                                  :parameters => {'calendarId' => 'primary'},
                                  :body => JSON.dump(event),
                                  :headers => {'Content-Type' => 'application/json'})

        json = JSON.parse(result.body)
        p json
        return json["summary"] == arg[:subject]
      rescue Exception => e
        p e
      end
    end
    false
  end

  ## 指定期間内の365のイベントリスト取得

  def get_office_tasks(timeMin, timeMax)
    p "365 CAL: " + timeMin.to_s + " -> " + timeMax.to_s
    uri = "https://outlook.office365.com/api/v1.0/me/calendarview?" \
          "startDatetime=#{timeMin.to_s}&endDateTime=#{timeMax.to_s}"

    certs = [$OFFICE_ID, $OFFICE_PASS]
    json = JSON.parse(open(uri, {:http_basic_authentication => certs}).read)
    json["value"].map do |t|
      {
        :subject => t["Subject"],
        :start => Time.parse(t["Start"]).localtime.iso8601.to_s,
        :end => Time.parse(t["End"]).localtime.iso8601.to_s,
        :isAllDay => t["IsAllDay"]
      }
    end
  end

  def help
    msg = <<-"AAA"
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





