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
        private static readonly int[] s_containerPropertyIds = new int[]
        {
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
            UIA_ControlTypeIds.UIA_CalendarControlTypeId,
            UIA_ControlTypeIds.UIA_DocumentControlTypeId,
            UIA_ControlTypeIds.UIA_EditControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_TabControlTypeId,
            UIA_ControlTypeIds.UIA_SliderControlTypeId,
            UIA_ControlTypeIds.UIA_SpinnerControlTypeId,
        };

        private static readonly int[] s_containerItemPropertyIds = new int[]
        {
            UIA_PropertyIds.UIA_IsExpandCollapsePatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsGridItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsScrollItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSelectionItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSpreadsheetItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsTableItemPatternAvailablePropertyId
        };

        private static readonly int[] s_noFocusControlTypeIds = new int[]
        {
            UIA_ControlTypeIds.UIA_GroupControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_WindowControlTypeId
        };

        private readonly IUIAutomation _automation;

        private readonly IUIAutomationCondition _accessibleCondition;

        private readonly IUIAutomationCacheRequest _cacheRequest;

        private readonly UiAutomationEnumerator _enumerator;

        public UiAutomationHintProviderService()
        {
            _automation = new CUIAutomation();

            _accessibleCondition = _automation.CreateAndCondition(
                    _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsEnabledPropertyId, true),
                    _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsOffscreenPropertyId, false));

            _cacheRequest = _automation.CreateCacheRequest();
            foreach (var propertyId in s_containerPropertyIds)
            {
                _cacheRequest.AddProperty(propertyId);
            }
            foreach (var propertyId in s_containerItemPropertyIds)
            {
                _cacheRequest.AddProperty(propertyId);
            }
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsEnabledPropertyId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsOffscreenPropertyId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
            _cacheRequest.AddPattern(UIA_PatternIds.UIA_InvokePatternId);
            _cacheRequest.AddPattern(UIA_PatternIds.UIA_TogglePatternId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_ControlTypePropertyId);

            _enumerator = new UiAutomationEnumerator(this);
        }

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
            var elements = _enumerator.EnumElements(hWnd);

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
                // TODO: noncontainer
                // TODO: focus part of item?
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

        private class UiAutomationEnumerator
        {
            private readonly UiAutomationHintProviderService _parent;

            private List<IUIAutomationElement> _result;

            public UiAutomationEnumerator(UiAutomationHintProviderService parent)
            {
                _parent = parent;
            }

            /// <summary>
            /// Enumerates the automation elements from the given window
            /// </summary>
            /// <param name="hWnd">The window handle</param>
            /// <returns>All of the automation elements found</returns>
            public List<IUIAutomationElement> EnumElements(IntPtr hWnd)
            {
                _result = new List<IUIAutomationElement>();
                AddElements(_parent._automation.ElementFromHandleBuildCache(hWnd, _parent._cacheRequest));
                return _result;
            }
            
            private void AddElements(IUIAutomationElement element)
            {
                var children = element.FindAllBuildCache(TreeScope.TreeScope_Children, _parent._accessibleCondition, _parent._cacheRequest);
                for (var i = 0; i < children.Length; ++i)
                {
                    var childElement = children.GetElement(i);
                    _result.Add(childElement);

                    if (s_nonContainerControlTypeIds.Contains(childElement.CachedControlType) ||
                        s_containerPropertyIds.All(propertyId => !(bool)childElement.GetCachedPropertyValue(propertyId)))
                    {
                        // non-container
                        AddElements(childElement);
                    }
                    else
                    {
                        // container
                        new ContainerEnumerator(this).AddContainerItems(childElement);
                    }
                }
            }

            // Typical container structure (list or tree)
            //     (Header)
            //     item
            //         parts of item
            //     item
            //         item
            //             parts of item
            //     ....
            //     (Footer)
            // Bidirectional recursive enumeration
            // Example:
            //     container
            //         item1 @ level1
            //             item2 @ level2
            //                 item3 @ level3 (on-screen)
            private class ContainerEnumerator
            {
                private readonly UiAutomationEnumerator _parent;

                private readonly IUIAutomationTreeWalker _walker;

                public ContainerEnumerator(UiAutomationEnumerator parent)
                {
                    _parent = parent;
                    _walker = _parent._parent._automation.RawViewWalker;
                }

                public void AddContainerItems(IUIAutomationElement containerElement)
                {
                    AddContainerItems2(containerElement, 0);
                }

                private void AddContainerItems2(IUIAutomationElement itemElement, int itemLevel)
                {
                    var elementIds = new HashSet<string>();
                    Func<IUIAutomationElement, bool> addNewElement = elem => elementIds.Add(
                        string.Join(".", from id in (int[])elem.GetRuntimeId() select id.ToString()));
                    Func<IUIAutomationElement, bool> isItemElement = elem => s_containerItemPropertyIds.Any(
                        propertyId => (bool)elem.GetCachedPropertyValue(propertyId));

                    // Enumerate on-screen items

                    // TODO: isoffscreen is false for zero size bounding rectangle
                    // Recursively enumerate on-screen elements forwards
                    IUIAutomationElement itemFoundFront = null;
                    for (var curElement = _walker.GetFirstChildElementBuildCache(itemElement, _parent._parent._cacheRequest);
                        curElement != null && curElement.CachedIsOffscreen == 0 && addNewElement(curElement);
                        curElement = _walker.GetNextSiblingElementBuildCache(curElement, _parent._parent._cacheRequest))
                    {
                        if (isItemElement(curElement))
                        {
                            // Found an on-screen item
                            // Break!
                            itemFoundFront = curElement;
                            break;
                        }
                        else if (curElement.CachedIsEnabled != 0)
                        {
                            // Find non-item elements
                            _parent.AddElements(curElement);
                        }
                    }

                    // Recursively enumerate on-screen elements backwards
                    IUIAutomationElement itemFoundBack = null;
                    for (var curElement = _walker.GetLastChildElementBuildCache(itemElement, _parent._parent._cacheRequest);
                        curElement != null && curElement.CachedIsOffscreen == 0 && addNewElement(curElement);
                        curElement = _walker.GetPreviousSiblingElementBuildCache(curElement, _parent._parent._cacheRequest))
                    {
                        if (isItemElement(curElement))
                        {
                            // Found an on-screen item
                            // Break!
                            itemFoundBack = curElement;
                            break;
                        }
                        else if (curElement.CachedIsEnabled != 0)
                        {
                            // Find non-item elements
                            _parent.AddElements(curElement);
                        }
                    }

                    if (itemFoundFront != null && itemFoundBack != null)
                    {
                        // Found item in both
                        // Recursively Enumerate all!
                        var children = itemElement.FindAllBuildCache(TreeScope.TreeScope_Children, 
                            _parent._parent._automation.CreateTrueCondition() , _parent._parent._cacheRequest);
                        for (var i = 0; i < children.Length; ++i)
                        {
                            var curItem = children.GetElement(i);
                            _parent._result.Add(curItem);
                            AddContainerItems2(curItem, itemLevel + 1);
                        }
                    }
                    else if (itemFoundFront != null)
                    {
                        // Found item in forward phase
                        // Recursively Enumerate forward
                        for (var curItem = itemFoundFront;
                            curItem != null && curItem.CachedIsOffscreen == 0;
                            curItem = _walker.GetNextSiblingElementBuildCache(curItem, _parent._parent._cacheRequest))
                        {
                            _parent._result.Add(curItem);
                            AddContainerItems2(curItem, itemLevel + 1);
                        }
                    }
                    else if (itemLevel == 0 && itemFoundBack != null)
                    {
                        // Found item in backward phase
                        // Recursively Enumerate backward
                        IUIAutomationElement frontElement;
                        for (frontElement = itemFoundBack;
                            frontElement != null && frontElement.CachedIsOffscreen == 0;
                            frontElement = _walker.GetPreviousSiblingElementBuildCache(frontElement, _parent._parent._cacheRequest))
                        {
                            _parent._result.Add(frontElement);
                            AddContainerItems2(frontElement, itemLevel + 1);
                        }
                        // descendant of pre-item1 may be on-screen
                        if (frontElement != null)
                        {
                            var j = 1;
                            for (var cur2 = frontElement;
                                            cur2 != null;
                                            cur2 = _walker.GetLastChildElementBuildCache(cur2,
                                                    _parent._parent._cacheRequest), ++j)
                            {
                                if (cur2.CachedIsOffscreen == 0)
                                {
                                    //      Enumerate backward from item2
                                    for (var curElement = cur2;
                                        curElement != null && curElement.CachedIsOffscreen == 0;
                                        curElement = _walker.GetPreviousSiblingElementBuildCache(curElement,
                                                _parent._parent._cacheRequest))
                                    {
                                        _parent._result.Add(curElement);
                                        AddContainerItems2(curElement, j + 1);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (itemLevel == 0)
                        {
                            // Search screen diagonally to find an on-screen item, e.g. item3
                            List<IUIAutomationElement> foundElementAncestors = new List<IUIAutomationElement>();
                            var boundingRectangle = itemElement.CurrentBoundingRectangle;
                            double delta = 8.0 / (boundingRectangle.right - boundingRectangle.left);

                            for (double t = 0; t < 1.0; t += delta)
                            {
                                foundElementAncestors.Clear();

                                tagPOINT p;
                                p.x = (int)Math.Floor(
                                    boundingRectangle.left + t * (boundingRectangle.right - boundingRectangle.left));
                                p.y = (int)Math.Floor(
                                    boundingRectangle.top + t * (boundingRectangle.bottom - boundingRectangle.top));

                                var elem = _parent._parent._automation.ElementFromPointBuildCache(p, _parent._parent._cacheRequest);
                                if (elem != null)
                                {
                                    var isWindowElement = false;

                                    for (var curElement = elem;
                                        curElement != null;
                                        curElement = _walker.GetParentElementBuildCache(curElement, _parent._parent._cacheRequest))
                                    {
                                        if (_parent._parent._automation.CompareElements(curElement, itemElement) != 0)
                                        {
                                            isWindowElement = true;
                                            break;
                                        }
                                        else if (isItemElement(curElement))
                                        {
                                            foundElementAncestors.Add(curElement);
                                        }
                                    }

                                    if (isWindowElement && foundElementAncestors.Count > 0)
                                    {
                                        break;
                                    }
                                }
                            }

                            if (foundElementAncestors.Count == 0)
                            {
                                // Still found none
                                // Done!
                                return;
                            }

                            foundElementAncestors.Reverse();

                            var curLevel = 0;
                            foreach (var cur in foundElementAncestors)
                            {
                                ++curLevel;
                                if (cur.CachedIsOffscreen == 0)
                                {
                                    // If item1 is on-screen
                                    // Enumerate backward/forward from item1
                                    IUIAutomationElement frontElement;
                                    for (
                                        frontElement =
                                            _walker.GetPreviousSiblingElementBuildCache(cur,
                                                _parent._parent._cacheRequest);
                                        frontElement != null && frontElement.CachedIsOffscreen == 0;
                                        frontElement =
                                            _walker.GetPreviousSiblingElementBuildCache(frontElement,
                                                _parent._parent._cacheRequest))
                                    {
                                        // Enumerate from level2 (forward+backward, forward)
                                        _parent._result.Add(frontElement);
                                        AddContainerItems2(frontElement, curLevel + 1);
                                    }

                                    // descendant of pre-item1 may be on-screen
                                    // pre-foundElementAncestors[0] if item1 is off-screen
                                    // pre-frontElement if item1 is on-screen
                                    if ((curLevel == 1 && frontElement != null) || curLevel > 1)
                                    {
                                        var preItem = curLevel == 1
                                            ? frontElement
                                            : _walker.GetPreviousSiblingElementBuildCache(
                                                foundElementAncestors[0], _parent._parent._cacheRequest);
                                        var j = 1;
                                        // If previous item1 is off-screen
                                        //  If last item2 is off-screen
                                        //          ...
                                        for (var cur2 = preItem;
                                            cur2 != null;
                                            cur2 = _walker.GetLastChildElementBuildCache(cur2,
                                                _parent._parent._cacheRequest), ++j)
                                        {
                                            if (cur2.CachedIsOffscreen == 0)
                                            {
                                                //      Enumerate backward from item2
                                                for (var curElement = cur2;
                                                    curElement != null && curElement.CachedIsOffscreen == 0;
                                                    curElement = _walker.GetPreviousSiblingElementBuildCache(curElement,
                                                        _parent._parent._cacheRequest))
                                                {
                                                    _parent._result.Add(curElement);
                                                    AddContainerItems2(curElement, j + 1);
                                                }
                                                break;
                                            }
                                        }
                                    }

                                    for (var curElement = cur;
                                        curElement != null && curElement.CachedIsOffscreen == 0;
                                        curElement =
                                            _walker.GetNextSiblingElementBuildCache(curElement,
                                                _parent._parent._cacheRequest))
                                    {
                                        _parent._result.Add(curElement);
                                        AddContainerItems2(curElement, curLevel + 1);
                                    }

                                    break;
                                }
                                else
                                {
                                    // descendant of post-item1 may be on-screen
                                    for (var curElement = _walker.GetNextSiblingElementBuildCache(
                                            cur, _parent._parent._cacheRequest);
                                        curElement != null && curElement.CachedIsOffscreen == 0;
                                        curElement = _walker.GetNextSiblingElementBuildCache(
                                            curElement, _parent._parent._cacheRequest))
                                    {
                                        _parent._result.Add(curElement);
                                        AddContainerItems2(curElement, curLevel + 1);
                                    }
                                }
                            }
                        }
                    }
                    // TODO: enabled?
                }
            }


        }
    }
}
