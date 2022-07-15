'SJIS�ŕۑ����Ȃ��Ɠ��삵�Ȃ��̂Œ���
Option Explicit

'�N�����̃��b�Z�[�W
Dim startmsg
startmsg = "Outlook�֗\���ǉ����܂��B"
MsgBox startmsg

'�\��̖��́A���t��Dictionary���`(���̂�Key)
Dim DicCalendarItems, DicCalendarItemKey
Set DicCalendarItems = CreateObject("Scripting.Dictionary")
DicCalendarItems.Add 1, Array("���a�̓�","2022/4/29")
DicCalendarItems.Add 2, Array("���@�L�O��","2022/5/3")
DicCalendarItems.Add 3, Array("�݂ǂ�̓�","2022/5/4")
DicCalendarItems.Add 4, Array("���ǂ��̓�","2022/5/5")
DicCalendarItems.Add 5, Array("�C�̓�","2022/7/18")
DicCalendarItems.Add 6, Array("�R�̓�","2022/8/11")
DicCalendarItems.Add 7, Array("�h�V�̓�","2022/9/19")
DicCalendarItems.Add 8, Array("�H���̓�","2022/9/23")
DicCalendarItems.Add 9, Array("�X�|�[�c�̓�","2022/10/10")
DicCalendarItems.Add 10, Array("�����̓�","2022/11/3")
DicCalendarItems.Add 11, Array("�ΘJ���ӂ̓�","2022/11/23")
DicCalendarItems.Add 12, Array("�U�֋x��","2023/1/2")
DicCalendarItems.Add 13, Array("���l�̓�","2023/1/9")
DicCalendarItems.Add 14, Array("�����L�O�̓�","2023/2/11")
DicCalendarItems.Add 15, Array("�V�c�a����","2023/2/23")
DicCalendarItems.Add 16, Array("�t���̓�","2023/3/21")

'Outlook�֓o�^
Const OutLookFolderCalendar = 9 '�f�t�H���g�J�����_�[
Const OutLookAppointItem = 1 '�쐬�A�C�e���̎�ށF�\��
Const OutLookBusyStatusFree = 0 '�X�e�[�^�X�F�\��Ȃ�
Const ItemCategoryName = "Sample_Category" '���̃c�[���œo�^�����\��ɃJ�e�S�������Ă���
COnst CategoryColorRed = 1 '�J�e�S���̐F�F��

Dim OutLookApp: Set OutLookApp = CreateObject("Outlook.Application")
Dim NameSpace: Set NameSpace = OutLookApp.GetNamespace("MAPI")
Dim OutLookFolder: Set OutLookFolder = NameSpace.GetDefaultFolder(OutLookFolderCalendar)

'�J�e�S�����Ȃ���ΐV�K�ɍ쐬
If NameSpace.Categories.Item(ItemCategoryName) Is Nothing Then
    NameSpace.Categories.Add ItemCategoryName, CategoryColorRed
End If

'�\��\��Item��ǉ�
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

'�������̃��b�Z�[�W
Dim endmsg
endmsg = "Outlook�֗\���ǉ����܂����B"
MsgBox endmsg
