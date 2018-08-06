require 'rubygems'
require './sync_task.rb'

term_month = if ARGV[0] == "help" then
               help = true
               nil
             else
               help = false
               # default: 今日から1ヶ月だけ
               (ARGV[0] || 1).to_i
             end

p Time.now
x = MyGCal.new

if help then
  x.help
  exit 0
end

timeMax = Date.today >> term_month
timeMin = Date.today - 1 # 全日予定をとるため -1 する

## Officeのタスク一覧を取ってきて
office_tasks = x.get_office_tasks(timeMin, timeMax)

## Googleのタスク一覧を取ってきて
google_task = x.get_google_tasks(timeMin, timeMax)

if google_task then

  ## 比較してGoogleにINSERT
  x.sync(office_tasks)

else
  exit 1
end

exit 0
