Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = True
App.WindowState = "Maximized"' Maximize the QuickTest window
App.ActivateView "ExpertView"' Display the Expert View
App.open "c:\tests\Book_Flight_One_Code", True
App.Test.Run