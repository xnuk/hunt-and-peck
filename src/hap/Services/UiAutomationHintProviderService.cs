using hap.Extensions;
using hap.Models;
using hap.NativeMethods;
using hap.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using UIAutomationClient;

namespace hap.Services
{
    internal class UiAutomationHintProviderService : IHintProviderService, IDebugHintProviderService
    {
        private readonly IUIAutomation _automation = new CUIAutomation();

        private static readonly int[] s_containerPropertyIds = new int[]
        {
            UIA_PropertyIds.UIA_IsExpandCollapsePatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsMultipleViewPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSelectionPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSpreadsheetPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId
        };

        private static readonly int[] s_nonContainerControlTypeIds = new int[]
        {
            UIA_ControlTypeIds.UIA_AppBarControlTypeId,
            UIA_ControlTypeIds.UIA_ButtonControlTypeId,
            UIA_ControlTypeIds.UIA_CalendarControlTypeId,
            UIA_ControlTypeIds.UIA_DocumentControlTypeId,
            UIA_ControlTypeIds.UIA_MenuBarControlTypeId,
            UIA_ControlTypeIds.UIA_MenuItemControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_TabControlTypeId,
            UIA_ControlTypeIds.UIA_SliderControlTypeId,
            UIA_ControlTypeIds.UIA_SpinnerControlTypeId,
            UIA_ControlTypeIds.UIA_SplitButtonControlTypeId,
            UIA_ControlTypeIds.UIA_ToolBarControlTypeId,
            UIA_ControlTypeIds.UIA_GroupControlTypeId
        };

        private static readonly int[] s_noFocusControlTypeIds = new int[]
        {
            UIA_ControlTypeIds.UIA_GroupControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_WindowControlTypeId
        };

        public HintSession EnumHints()
        {
            var foregroundWindow = User32.GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                return null;
            }
            return EnumHints(foregroundWindow);
        }

        public HintSession EnumHints(IntPtr hWnd)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            var session = EnumWindowHints(hWnd, CreateHint);
            sw.Stop();

            Debug.WriteLine("Enumeration of hints took {0} ms", sw.ElapsedMilliseconds);
            return session;
        }

        public HintSession EnumDebugHints()
        {
            var foregroundWindow = User32.GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                return null;
            }
            return EnumDebugHints(foregroundWindow);
        }

        public HintSession EnumDebugHints(IntPtr hWnd)
        {
            return EnumWindowHints(hWnd, CreateDebugHint);
        }

        /// <summary>
        /// Enumerates all the hints from the given window
        /// </summary>
        /// <param name="hWnd">The window to get hints from</param>
        /// <param name="hintFactory">The factory to use to create each hint in the session</param>
        /// <returns>A hint session</returns>
        private HintSession EnumWindowHints(IntPtr hWnd, Func<IntPtr, Rect, IUIAutomationElement, Hint> hintFactory)
        {
            var result = new List<Hint>();
            var elements = EnumElements(hWnd);

            // Window bounds
            var rawWindowBounds = new RECT();
            User32.GetWindowRect(hWnd, ref rawWindowBounds);
            Rect windowBounds = rawWindowBounds;

            foreach (var element in elements)
            {
                var boundingRectObject = element.CachedBoundingRectangle;
                if ((boundingRectObject.right > boundingRectObject.left) && (boundingRectObject.bottom > boundingRectObject.top))
                {
                    var niceRect = new Rect(new Point(boundingRectObject.left, boundingRectObject.top), new Point(boundingRectObject.right, boundingRectObject.bottom));
                    // Convert the bounding rect to logical coords
                    var logicalRect = niceRect.PhysicalToLogicalRect(hWnd);
                    if (!logicalRect.IsEmpty)
                    {
                        var windowCoords = niceRect.ScreenToWindowCoordinates(windowBounds);
                        var hint = hintFactory(hWnd, windowCoords, element);
                        if (hint != null)
                        {
                            result.Add(hint);
                        }
                    }
                }
            }

            return new HintSession
            {
                Hints = result,
                OwningWindow = hWnd,
                OwningWindowBounds = windowBounds,
            };
        }

        /// <summary>
        /// Enumerates the automation elements from the given window
        /// </summary>
        /// <param name="hWnd">The window handle</param>
        /// <returns>All of the automation elements found</returns>
        private List<IUIAutomationElement> EnumElements(IntPtr hWnd)
        {
            var result = new List<IUIAutomationElement>();

            var cacheRequest = _automation.CreateCacheRequest();
            foreach (var propertyId in s_containerPropertyIds)
            {
                cacheRequest.AddProperty(propertyId);
            }
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
            cacheRequest.AddPattern(UIA_PatternIds.UIA_InvokePatternId);
            cacheRequest.AddPattern(UIA_PatternIds.UIA_TogglePatternId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId);
            cacheRequest.AddProperty(UIA_PropertyIds.UIA_ControlTypePropertyId);

            var conditionControlView = _automation.ControlViewCondition;
            var conditionEnabled = _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsEnabledPropertyId, true);
            var enabledControlCondition = _automation.CreateAndCondition(conditionControlView, conditionEnabled);

            var conditionOnScreen = _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsOffscreenPropertyId, false);
            var condition = _automation.CreateAndCondition(enabledControlCondition, conditionOnScreen);

            var elementStack = new Stack<IUIAutomationElement>();
            elementStack.Push(_automation.ElementFromHandle(hWnd));

            while (elementStack.Count > 0)
            {
                var element = elementStack.Pop();
                var children = element.FindAllBuildCache(TreeScope.TreeScope_Children, condition, cacheRequest);

                for (var i = 0; i < children.Length; ++i)
                {
                    var childElement = children.GetElement(i);
                    
                    if (s_nonContainerControlTypeIds.Contains(childElement.CachedControlType) ||
                        s_containerPropertyIds.All(propertyId => !(bool) childElement.GetCachedPropertyValue(propertyId)))
                    {
                        // non-container
                        elementStack.Push(childElement);
                    }

                    result.Add(childElement);
                }
            }

            return result;
        }

        /// <summary>
        /// Creates a UI Automation element from the given automation element
        /// </summary>
        /// <param name="owningWindow">The owning window</param>
        /// <param name="hintBounds">The hint bounds</param>
        /// <param name="automationElement">The associated automation element</param>
        /// <returns>The created hint, else null if the hint could not be created</returns>
        private Hint CreateHint(IntPtr owningWindow, Rect hintBounds, IUIAutomationElement automationElement)
        {
            try
            {
                var invokePattern = (IUIAutomationInvokePattern) automationElement.GetCachedPattern(
                    UIA_PatternIds.UIA_InvokePatternId);
                if (invokePattern != null)
                {
                    return new UiAutomationInvokeHint(owningWindow, invokePattern, hintBounds);
                }

                var togglePattern = (IUIAutomationTogglePattern) automationElement.GetCachedPattern(
                    UIA_PatternIds.UIA_TogglePatternId);
                if (togglePattern != null)
                {
                    return new UiAutomationToggleHint(owningWindow, togglePattern, hintBounds);
                }

                var isFocusable = automationElement.CachedIsKeyboardFocusable != 0;
                if (isFocusable && !s_noFocusControlTypeIds.Contains(automationElement.CachedControlType))
                {
                    return new UiAutomationFocusHint(owningWindow, automationElement, hintBounds);
                }

                var isContainer = s_containerPropertyIds.Any(
                    propertyId => (bool) automationElement.GetCachedPropertyValue(propertyId));
                if (isContainer)
                {
                    return new UiAutomationFocusHint(owningWindow, automationElement, hintBounds);
                }

                return null;
            }
            catch (Exception)
            {
                // May have gone
                return null;
            }
        }

        /// <summary>
        /// Creates a debug hint
        /// </summary>
        /// <param name="owningWindow">The window that owns the hint</param>
        /// <param name="hintBounds">The hint bounds</param>
        /// <param name="automationElement">The automation element</param>
        /// <returns>A debug hint</returns>
        private DebugHint CreateDebugHint(IntPtr owningWindow, Rect hintBounds, IUIAutomationElement automationElement)
        {
            // Enumerate all possible patterns. Note that the performance of this is *very* bad -- hence debug only.
            var programmaticNames = new List<string>();

            foreach (var pn in UiAutomationPatternIds.PatternNames)
            {
                try
                {
                    var pattern = automationElement.GetCurrentPattern(pn.Key);
                    if(pattern != null)
                    {
                        programmaticNames.Add(pn.Value);
                    }
                }
                catch (Exception)
                {
                }
            }

            if (programmaticNames.Any())
            {
                return new DebugHint(owningWindow, hintBounds, programmaticNames.ToList());
            }

            return null;
        }
    }
}
