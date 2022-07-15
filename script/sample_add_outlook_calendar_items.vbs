'SJISで保存しないと動作しないので注意
Option Explicit

'起動時のメッセージ
Dim startmsg
startmsg = "Outlookへ予定を追加します。"
MsgBox startmsg

'予定の名称、日付のDictionaryを定義(名称がKey)
Dim DicCalendarItems, DicCalendarItemKey
Set DicCalendarItems = CreateObject("Scripting.Dictionary")
DicCalendarItems.Add 1, Array("昭和の日","2022/4/29")
DicCalendarItems.Add 2, Array("憲法記念日","2022/5/3")
DicCalendarItems.Add 3, Array("みどりの日","2022/5/4")
DicCalendarItems.Add 4, Array("こどもの日","2022/5/5")
DicCalendarItems.Add 5, Array("海の日","2022/7/18")
DicCalendarItems.Add 6, Array("山の日","2022/8/11")
DicCalendarItems.Add 7, Array("敬老の日","2022/9/19")
DicCalendarItems.Add 8, Array("秋分の日","2022/9/23")
DicCalendarItems.Add 9, Array("スポーツの日","2022/10/10")
DicCalendarItems.Add 10, Array("文化の日","2022/11/3")
DicCalendarItems.Add 11, Array("勤労感謝の日","2022/11/23")
DicCalendarItems.Add 12, Array("振替休日","2023/1/2")
DicCalendarItems.Add 13, Array("成人の日","2023/1/9")
DicCalendarItems.Add 14, Array("建国記念の日","2023/2/11")
DicCalendarItems.Add 15, Array("天皇誕生日","2023/2/23")
DicCalendarItems.Add 16, Array("春分の日","2023/3/21")

'Outlookへ登録
Const OutLookFolderCalendar = 9 'デフォルトカレンダー
Const OutLookAppointItem = 1 '作成アイテムの種類：予定
Const OutLookBusyStatusFree = 0 'ステータス：予定なし
Const ItemCategoryName = "Sample_Category" 'このツールで登録した予定にカテゴリをつけておく
COnst CategoryColorRed = 1 'カテゴリの色：赤

Dim OutLookApp: Set OutLookApp = CreateObject("Outlook.Application")
Dim NameSpace: Set NameSpace = OutLookApp.GetNamespace("MAPI")
Dim OutLookFolder: Set OutLookFolder = NameSpace.GetDefaultFolder(OutLookFolderCalendar)

'カテゴリがなければ新規に作成
If NameSpace.Categories.Item(ItemCategoryName) Is Nothing Then
    NameSpace.Categories.Add ItemCategoryName, CategoryColorRed
End If

'予定表にItemを追加
For Each DicCalendarItemKey In DicCalendarItems.Keys
    Dim NewOutLookItem 'As Outlook.AppointmentItem
    Set NewOutLookItem = OutLookApp.CreateItem(OutLookAppointItem)
    With NewOutLookItem
        .Subject = DicCalendarItems(DicCalendarItemKey)(0)
        .Start = DicCalendarItems(DicCalendarItemKey)(1)
        .AllDayEvent = True
        .BusyStatus = OutLookBusyStatusFree
        .Categories = ItemCategoryName
        .Save
        .Close 0
    End With
Next

'完了時のメッセージ
Dim endmsg
endmsg = "Outlookへ予定を追加しました。"
MsgBox endmsg
