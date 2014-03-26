VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlExpressionsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SQLFieldExpressionTest()

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    
    objFieldExpression.FieldName = "Field1"
    
    Assert.That GetExpression(objFieldExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("[Field1]")

End Sub

Public Sub SQLValueExpressionTest()

    Dim objExpression As SQLValueExpression
    Set objExpression = New SQLValueExpression
    
    objExpression.Value = "123"
    
    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("'123'")

End Sub

Public Sub SQLValueNullExpressionTest()

    Dim objExpression As SQLValueExpression
    Set objExpression = New SQLValueExpression
    
    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("NULL")

End Sub

Public Sub ArithmeticExpressionTest()

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    objFieldExpression.FieldName = "Field1"

    Dim objValueExpression As SQLValueExpression
    Set objValueExpression = New SQLValueExpression
    objValueExpression.Value = 1

    Dim objExpression As SQLArithmeticExpression
    Set objExpression = New SQLArithmeticExpression
    
    Set objExpression.Left = objFieldExpression
    objExpression.Operator = dboArithmeticAdd
    Set objExpression.Right = objValueExpression
    
    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("([Field1] + 1)")

End Sub

Public Sub BitwiseExpressionTest()

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    objFieldExpression.FieldName = "Field1"

    Dim objValueExpression As SQLValueExpression
    Set objValueExpression = New SQLValueExpression
    objValueExpression.Value = 1

    Dim objExpression As SQLBitwiseExpression
    Set objExpression = New SQLBitwiseExpression
    
    Set objExpression.Left = objFieldExpression
    objExpression.Operator = dboBitwiseAnd
    Set objExpression.Right = objValueExpression
    
    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("([Field1] & 1)")

End Sub

Public Sub ComparisonExpressionTest()

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    objFieldExpression.FieldName = "Field1"

    Dim objValueExpression As SQLValueExpression
    Set objValueExpression = New SQLValueExpression
    objValueExpression.Value = 1

    Dim objExpression As SQLComparisonExpression
    Set objExpression = New SQLComparisonExpression
    
    Set objExpression.Left = objFieldExpression
    objExpression.Compare = dboComparisonGreaterThan
    Set objExpression.Right = objValueExpression
    
    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("([Field1] > 1)")

End Sub

Public Sub ConditionsCompareExpression()

    Dim objSelect As SQLSelect
    Set objSelect = New SQLSelect
    objSelect.ConnectionType = dboConnectionTypeSQLServer
    objSelect.Tables.Add "Table1"

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    objFieldExpression.FieldName = "Field1"

    Dim objValueExpression As SQLValueExpression
    Set objValueExpression = New SQLValueExpression
    objValueExpression.Value = 1

    objSelect.Where.AddCompareExpression objFieldExpression, dboComparisonEqualTo, objValueExpression
    
    Assert.That objSelect.SQL, Iz.EqualTo("SELECT * FROM [Table1] WHERE ([Field1] = 1)")

End Sub

Public Sub SubExpressions()

    Dim objFieldExpression As SqlFieldExpression
    Set objFieldExpression = New SqlFieldExpression
    objFieldExpression.FieldName = "Field1"

    Dim objValueExpression As SQLValueExpression
    Set objValueExpression = New SQLValueExpression
    objValueExpression.Value = 1

    Dim objSubExpression As SQLArithmeticExpression
    Set objSubExpression = New SQLArithmeticExpression
    Set objSubExpression.Left = objFieldExpression
    objSubExpression.Operator = dboArithmeticAdd
    Set objSubExpression.Right = objValueExpression

    Dim objExpression As SQLComparisonExpression
    Set objExpression = New SQLComparisonExpression
    Set objExpression.Left = objFieldExpression
    objExpression.Compare = dboComparisonEqualTo
    Set objExpression.Right = objSubExpression

    Assert.That GetExpression(objExpression).SQL(dboConnectionTypeSQLServer), Iz.EqualTo("([Field1] = ([Field1] + 1))")

End Sub

Private Function GetExpression(ByVal objExpression As ISQLExpression) As ISQLExpression

    Set GetExpression = objExpression

End Function
