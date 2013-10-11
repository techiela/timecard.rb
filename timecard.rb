require 'win32/eventlog'
require 'date'
require 'rubygems'
require 'active_support/time'
include Win32

class Timecard

  #
  # コンストラクタ
  #
  def initialize
    self.initializeEventLogContainer
  end

  #
  # イベントログ格納配列を初期化する
  #
  def initializeEventLogContainer
    @result = {}
    today = Date.today;

    # 先月の１日から今月の末日まで
    (today.months_ago(1).beginning_of_month..today.end_of_month).each do |day|
      @result[day.strftime('%Y-%m-%d')] = []
    end
  end

  #
  # 起動メソッド
  #
  def main

    self.read

    self.buffer

    self.put

  end

  #
  # windowsのイベントログ(system)を読み込む
  #
  def read

    prevMonth = Date.today.months_ago(1).month
    currentMonth = Date.today.month

    # systemログを開く
    handle = EventLog.open('system')

    i = 0
    # 直近の10000行を読み込むようにオフセットを計算する
    handle.read(EventLog::SEQUENTIAL_READ | EventLog::BACKWARDS_READ) do |log|
      if 10000 <= i += 1
        break
      end
      # 先月・今月のデータのみを処理対象とする
      if log.time_written.month != currentMonth &&
         log.time_written.month != prevMonth
         next
      end
      @result[log.time_written.strftime('%Y-%m-%d')]
        .push(log.time_written.strftime('%Y-%m-%d %H:%M'))
    end
    handle.close
  end

  #
  # 計算結果をバッファリングする
  #
  def buffer
    @buffer = "date\tstartTime\tendTime\trest\tactualWorkedHours\n"
    @result.each do |key, value|

      # イベントのない日をスキップ
      if value.length == 0
        @buffer += sprintf("%s\n", key)
        next
      end

      # PC起動時刻を取得
      boot = self.round(value.min)
      
      # PC起動時刻を09:30始業に合わせる
      if boot.strftime('%H%M') < '0930'
        boot = DateTime.parse(boot.strftime('%Y-%m-%d ') + '09:30');
      end

      # PCシャットダウン(≒その日の最終イベント)時刻取得
      shutdown = self.round(value.max)

      # 実働時間計算
      diff = (((shutdown - boot) * 24 * 60).to_i); # diff minutes

      # 休憩時間
      rest = 1
      if '1400' <= boot.strftime('%H%M') ||
        shutdown.strftime('%H%M') <= '1400'
        # 午後休 or 午後出社
        rest = 0
      end

      # 出力バッファに１行追加
      @buffer += sprintf("%s\t%s\t%s\t%d:00\t%d.%02d\n",
        key,
        boot.strftime('%H:%M'),
        shutdown.strftime('%H:%M'),
        rest,
        diff / 60 - rest,
        diff % 60 / 15 * 25)
    end
  end

  #
  # 時刻を15分刻みにして返す(時刻の分のところを 00, 15, 30, 45 のみにする)
  #
  def round(value)
    d = DateTime.parse(value)
    odd = d.min % 15
    if odd <= 7
      return d - Rational(odd, 24 * 60)
    else
      return d + Rational(15 - odd, 24 * 60)
    end
  end

  #
  # 計算結果を出力する
  #
  def put
    print @buffer
    File.write('timecard.tsv', @buffer)
  end

end

# run
timecard = Timecard.new
timecard.main
