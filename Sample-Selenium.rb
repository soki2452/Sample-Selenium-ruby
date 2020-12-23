# encoding: Windows-31J
require 'selenium-webdriver'
require 'win32ole'
require 'fileutils'
require 'open-uri'
 
#エレメント検索Key
KEY_ITEM_SHOP = 'h5.s-line-clamp-1'
KEY_ITEM_NAME = 'span.a-size-base-plus.a-color-base.a-text-normal'
KEY_ITEM_PRICE = 'a-price-whole'
KEY_ITEM_IMG = 'img.s-image'
KEY_ITEM_STAR = 'span.a-icon-alt'


class Scripting_Amazon
    #要素がない場合、例外で処理が終わってしまうので、例外の場合、falseを返却して処理を続ける
    def is_element_present?(el, how, what)
        @session.manage.timeouts.implicit_wait = 0 #ロードされるまで待機時間を変更（このメソッドだけ）
        el.find_element(how, what)
        true
    rescue Selenium::WebDriver::Error::NoSuchElementError
        false
    end

    #メインメソッド
    def main
        fso = WIN32OLE.new("Scripting.FileSystemObject")
        #画像用ダウンロード先のフォルダがあるか存在確認、存在しなければ作成する
        img_path = 'c:\\temp\\image\\work_img.jpg'  #画像ファイルの保存先（ファイル名は固定）
        FileUtils.touch(img_path) unless FileTest.exists?(img_path)

        #Excel（メニュー、原紙）をOpen
        excel = WIN32OLE.new("Excel.Application")
        excel.visible = true                    #アプリケーションを表示する
        #WIN32OLEを使用してファイルをOpenする場合、フルパスの指定が必要となる
        wb = excel.Workbooks.Open(:FileName => fso.GetAbsolutePathName("Amazon.xlsx"), :ReadOnly => true)
        #メニューから検索条件を取得
        ws = wb.Worksheets("メニュー")
        amazon_url = ws.Range("F3").Value
        key_word = ws.Range("B3").Value
        sort_index = ws.Range("C3").Value
        max_count = ws.Range("D3").Value + 1    #セルの位置が2行目から開始しているため
    
        #検索結果シートの作成
        wb.Worksheets("検索結果").Copy
        wb.close                                #元のExcelは不要のためClose
        wb = excel.ActiveWorkbook               #対象を新規に作成したExcelとする
        ws = wb.Worksheets("検索結果")

        #Chromeを起動し、AmazonのTopページを表示する
        @session = Selenium::WebDriver.for :chrome
        @session.manage.timeouts.implicit_wait = 600 #ロードされるまで待機
        @session.get amazon_url

        #タイトルにAmazonが含まれていれば商品ページと判定
        if @session.title.index("Amazon").nil?
            puts '商品ページの取得ができませんでした。'
        #    End
        end

        #検索条件にキーワードを入力し、クリック
        @session.find_element(:id, "twotabsearchtextbox").send_keys key_word
        @session.find_element(:class, "nav-right").click

        #並び替えセレクトボックスの選択肢を表示する
        @session.find_element(:id, "a-autoid-0-announce").click
        #選択肢の中から、セルに入力された項目を選択
        @session.find_element(:link_text, sort_index).click
        @session.manage.timeouts.implicit_wait = 600 #ロードされるまで待機（これをやらないと後続の処理でnotfoundになる）

        row = 2
        i = 3
        shapes_ount = 0
        until row >= max_count do
            #商品検索結果のページが3～54配列の52件となっているため
            for i in 3..54 do
                element = @session.find_element(:xpath, "//*[@id=\"search\"]/div[1]/div[2]/div/span[3]/div[2]/div[#{i}]")
                #商品名
                ws.cells.item(row, 4).value = element.find_element(:css, KEY_ITEM_NAME).text
                #商品画像添付
                if is_element_present?(element, :css, KEY_ITEM_IMG)
                    #puts element.find_element(:css, KEY_ITEM_IMG).attribute("src")
                    #ファイルのダウンロード実行
                    url = element.find_element(:css, KEY_ITEM_IMG).attribute("src")
                    #file = img_path
                    URI.open(url) { |image|
                        File.open(img_path,"wb") do |file|
                            file.puts image.read
                        end
                    }
                    #ダウンロードしたファイルをセルに挿入
                    ws.range("B#{row}").select
                    ws.Pictures.Insert(img_path)
                    shapes_ount += 1
                else
                    ws.cells.item(row, 2).value = 'N/A'
                end
                #出品者情報
                if is_element_present?(element, :css, KEY_ITEM_SHOP)
                    ws.cells.item(row, 3).value = element.find_element(:css, KEY_ITEM_SHOP).text
                else
                    ws.cells.item(row, 3).value = '-'
                end
                #価格
                if is_element_present?(element, :class, KEY_ITEM_PRICE)
                    ws.cells.item(row, 5).value = element.find_element(:class, KEY_ITEM_PRICE).text
                else
                    ws.cells.item(row, 5).value = '-'
                end
                #商品URL
                ws.cells.item(row, 6).value = 'https://www.amazon.co.jp/dp/' + element.attribute("data-asin")
                ws.Hyperlinks.add("Anchor"=>ws.cells(row, 6), "Address"=>ws.cells.item(row, 6).value)
                #評価
                if is_element_present?(element, :css, KEY_ITEM_STAR)
                    ws.cells.item(row, 7).value = element.find_element(:css, KEY_ITEM_STAR).attribute("innerHTML")
                else
                    ws.cells.item(row, 7).value = 'N/A'
                end
                #要求された件数を満たした場合、処理を終了する
                if row >= max_count
                    break
                end
                row = row + 1
            end
            #要求された件数に満たない場合、次のページを表示し、処理を繰り返す
            if row < max_count 
                if is_element_present?(@session, :css, "li.a-last")
                    @session.find_element(:css, "li.a-last").click
                end
            end
        end

        #貼り付けた画像のサイズをセル内に収める
        for i in 1..shapes_ount do
            ws.Shapes(i).IncrementLeft(6)
            ws.Shapes(i).IncrementTop(6)
            if ws.Shapes(i).Height > 90 
                ws.Shapes(i).Height = 85
            end
            if ws.Shapes(i).Width > 200 
                ws.Shapes(i).Width = 200
            end
        end

        #ドライバーを閉じる
        excel.Quit
        @session.quit
    end
end

s = Scripting_Amazon.new
s.main
