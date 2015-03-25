Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports ClassLibrary1



'''<summary>
'''Class1Test のテスト クラスです。すべての
'''Class1Test 単体テストをここに含めます
'''</summary>
<TestClass()> _
Public Class Class1Test


    Private testContextInstance As TestContext

    '''<summary>
    '''現在のテストの実行についての情報および機能を
    '''提供するテスト コンテキストを取得または設定します。
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "追加のテスト属性"
    '
    'テストを作成するときに、次の追加属性を使用することができます:
    '
    'クラスの最初のテストを実行する前にコードを実行するには、ClassInitialize を使用
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    'クラスのすべてのテストを実行した後にコードを実行するには、ClassCleanup を使用
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    '各テストを実行する前にコードを実行するには、TestInitialize を使用
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    '各テストを実行した後にコードを実行するには、TestCleanup を使用
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region

    Private Sub Sub_DoubleValueTest()
        Dim input
        Dim output
        Dim err_flg As Boolean

        Try
            input = TestContext.DataRow("入力値")
            output = TestContext.DataRow("出力値")
        Catch
            err_flg = True
        End Try

        If String.Equals(TestContext.DataRow("エラー発生？"), "TRUE") Then
            Try
                Assert.AreEqual(Of Byte)(ClassLibrary1.Class1.DoubleValue(input), output)
            Catch
                err_flg = True
            End Try

            If err_flg = False Then
                Assert.Fail()
            End If
        Else
            Assert.AreEqual(Of Byte)(ClassLibrary1.Class1.DoubleValue(input), output)
        End If
    End Sub

    '''<summary>
    '''DoubleValue のテスト
    '''</summary>
    <DeploymentItem("TestProject1\CSVData.csv")> <DeploymentItem("TestProject1\NG_CSVData.csv")> <Description("テストが失敗するデータを使用しています")> <DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\NG_CSVData.csv", "NG_CSVData#csv", DataAccessMethod.Sequential)> <TestMethod()> _
    Public Sub NG_DoubleValueTest()
        Sub_DoubleValueTest()
    End Sub

    '''<summary>
    '''DoubleValue のテスト
    '''</summary>
    <Description("テストが成功するデータ")> <DeploymentItem("TestProject1\CSVData.csv")> <DeploymentItem("TestProject1\OK_CSVData.csv")> <DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\OK_CSVData.csv", "OK_CSVData#csv", DataAccessMethod.Sequential)> <TestMethod()> _
    Public Sub OK_DoubleValueTest()
        Sub_DoubleValueTest
    End Sub

End Class
