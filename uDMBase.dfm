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
end
