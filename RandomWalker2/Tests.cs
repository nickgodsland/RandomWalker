using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using NUnit.Framework;

namespace RandomWalker2
{
   [TestFixture]
   public class Tests
   {
      private Process _process;
      private AutomationElementCollection _automationElements;

      private readonly string[] _elementBlackList =
      {
         "Minimize", "Maximize", "Close", "System", "Calculator", "Application", "System Menu Bar"
      };

      private enum MouseEvent
      {
         MouseEventLeftDown = 0x02,
         MouseEventLeftUp = 0x04
      }

      [TearDown]
      public void TearDown()
      {
         Automation.RemoveAllEventHandlers();
      }

      [Test]
      public void Test()
      {
         StartProces();

         var mainWindowAutomationElement = AutomationElement.FromHandle(_process.MainWindowHandle);

         _automationElements = mainWindowAutomationElement.FindAll(TreeScope.Subtree, Condition.TrueCondition);

         StructureChangedEventHandler structureChangedEventHandler = OnStructureChange;
         Automation.AddStructureChangedEventHandler(mainWindowAutomationElement, TreeScope.Subtree,
            structureChangedEventHandler);

         var random = new Random();

         while (!_process.HasExited && ProcessIsActiveWindow())
         {
            var randomAutomationElement = _automationElements[random.Next(_automationElements.Count - 1)];

            try
            {
               Console.WriteLine($"Name: {randomAutomationElement.Current.Name} ClassName: {randomAutomationElement.Current.ClassName}");

               if (CheckElementIsInvalid(randomAutomationElement))
                  continue;

               var location = randomAutomationElement.Current.BoundingRectangle.Location;

               SetCursorPos((int) location.X, (int) location.Y);

               Click((int) MouseEvent.MouseEventLeftDown, 0, 0, 0, 0);
               Click((int) MouseEvent.MouseEventLeftUp, 0, 0, 0, 0);
            }
            catch (ElementNotAvailableException)
            {
               Console.WriteLine("ElementNotAvailableException");
            }
         }
      }

      private void OnStructureChange(object sender, StructureChangedEventArgs e)
      {
         Console.WriteLine($"OnStructureChange {e.StructureChangeType}, {e.EventId}, {e.GetRuntimeId()}");

         var automationElement = AutomationElement.FromHandle(_process.MainWindowHandle);
         _automationElements = automationElement.FindAll(TreeScope.Subtree, Condition.TrueCondition);
      }

      private bool CheckElementIsInvalid(AutomationElement randomAutomationElement)
      {
         return randomAutomationElement.Current.BoundingRectangle.IsEmpty ||
                randomAutomationElement.Current.IsOffscreen ||
                _elementBlackList.Contains(randomAutomationElement.Current.Name);
      }

      private void StartProces(string name = "calc.exe")
      {
         _process = Process.Start(name);
         while (_process.MainWindowHandle == IntPtr.Zero)
         {
         }
      }

      private bool ProcessIsActiveWindow()
      {
         var foregroundWindowHandle = GetForegroundWindow();
         var isActive = _process.MainWindowHandle == foregroundWindowHandle || foregroundWindowHandle == IntPtr.Zero;

         if (!isActive)
         {
            Console.WriteLine($"Process no longer has active window");
         }

         return isActive;
      }

      [DllImport("user32.dll")]
      private static extern bool SetCursorPos(int positionX, int positionY);

      [DllImport("user32.dll")]
      private static extern IntPtr GetForegroundWindow();

      [DllImport("User32.dll", EntryPoint = "mouse_event")]
      private static extern void Click(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
   }
}
