object DMbase: TDMbase
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 323
  Width = 386
  object MainConnection: TFDConnection
    Params.Strings = (
      'LockingMode=Normal'
      'JournalMode=WAL'
      'DateTimeFormat=DateTime'
      'Extensions=True'
      'DriverID=SQLite')
    Left = 32
    Top = 24
  end
  object Exams: TFDQuery
    Connection = MainConnection
    SQL.Strings = (
      'SELECT * from [TmsExams.ADT]')
    Left = 32
    Top = 104
  end
  object Items: TFDQuery
    Connection = MainConnection
    FetchOptions.AssignedValues = [evUnidirectional, evCursorKind]
    FetchOptions.Unidirectional = True
    SQL.Strings = (
      'SELECT * FROM [ITEMS.ADT]')
    Left = 128
    Top = 104
  end
  object DeleteItems: TFDQuery
    Connection = MainConnection
    SQL.Strings = (
      'DELETE FROM [ITEMS.ADT]'
      'WHERE Item_ID = :Param')
    Left = 136
    Top = 176
    ParamData = <
      item
        Name = 'PARAM'
        DataType = ftInteger
        ParamType = ptInput
        Value = Null
      end>
  end
end
