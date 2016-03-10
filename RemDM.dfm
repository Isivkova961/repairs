object dmRem: TdmRem
  OldCreateOrder = False
  Left = 344
  Top = 141
  Height = 150
  Width = 215
  object dsRem: TDataSource
    DataSet = adotRem
    Left = 112
    Top = 8
  end
  object adocRem: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=data.mdb;Persist S' +
      'ecurity Info=False'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 16
    Top = 8
  end
  object adotRem: TADOTable
    Active = True
    Connection = adocRem
    CursorType = ctStatic
    TableName = 'remont'
    Left = 64
    Top = 8
  end
end
