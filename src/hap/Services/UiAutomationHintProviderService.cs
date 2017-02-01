using hap.Extensions;
using hap.Models;
using hap.NativeMethods;
using hap.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using Caliburn.Micro;
using UIAutomationClient;

namespace hap.Services
{
    internal class UiAutomationHintProviderService : IHintProviderService, IDebugHintProviderService
    {
        // TODO: wrap in method
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

        private readonly IUIAutomationTreeWalker _treeWalker;
        
        private readonly IUIAutomationCondition _accessibleCondition;

        private readonly IUIAutomationCacheRequest _cacheRequest;

        public UiAutomationHintProviderService()
        {
            _automation = new CUIAutomation();

            _treeWalker = _automation.RawViewWalker;

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

            Action<IUIAutomationElement> addElements = null;
            Action<IUIAutomationElement, bool> addContainerItems = null;

            addElements = element =>
            {
                result.Add(element);

                if (s_nonContainerControlTypeIds.Contains(element.CachedControlType) ||
                    s_containerPropertyIds.All(propertyId => !(bool) element.GetCachedPropertyValue(propertyId)))
                {
                    GetElements(element.FindAllBuildCache(TreeScope.TreeScope_Children, _accessibleCondition, _cacheRequest))
                        .Apply(child => addElements(child));
                }
                else
                {
                    addContainerItems(element, false);
                }
            };

            // Enumerate on-screen items
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
            Func<IUIAutomationElement, bool> isItemElement = elem => s_containerItemPropertyIds.Any(
                propertyId => (bool)elem.GetCachedPropertyValue(propertyId));

            addContainerItems = (itemElement, itemLevel) =>
            {
                if (itemLevel && itemElement.CachedIsEnabled != 0)
                {
                    result.Add(itemElement);
                }

                var elementIds = new HashSet<string>();
                Func<IUIAutomationElement, bool> addNewElement = elem => elementIds.Add(
                    string.Join(".", from id in (int[])elem.GetRuntimeId() select id.ToString()));

                // Recursively enumerate on-screen elements forwards
                IUIAutomationElement itemFoundFront = null;
                GetNextSiblingElementsBuildCache(_treeWalker.GetFirstChildElementBuildCache(
                        itemElement, _cacheRequest))
                    .TakeWhile(e => e.CachedIsOffscreen == 0 && addNewElement(e))
                    .TakeWhile(curElement =>
                    {
                        if (!isItemElement(curElement))
                        {
                            return true;
                        }
                        // Found an on-screen item
                        // Break!
                        itemFoundFront = curElement;
                        return false;
                    })
                    .Apply(curElement =>
                    {
                        if (curElement.CachedIsEnabled != 0)
                        {
                            // Find non-item elements
                            addElements(curElement);
                        }
                    });

                // Recursively enumerate on-screen elements backwards
                IUIAutomationElement itemFoundBack = null;
                GetPreviousSiblingElementsBuildCache(_treeWalker.GetLastChildElementBuildCache(
                        itemElement, _cacheRequest))
                    .TakeWhile(e => e.CachedIsOffscreen == 0 && addNewElement(e))
                    .TakeWhile(curElement =>
                    {
                        if (!isItemElement(curElement))
                        {
                            return true;
                        }
                        // Found an on-screen item
                        // Break!
                        itemFoundBack = curElement;
                        return false;
                    })
                    .Apply(curElement =>
                    {
                        if (curElement.CachedIsEnabled != 0)
                        {
                            // Find non-item elements
                            addElements(curElement);
                        }
                    });

                if (itemFoundFront != null && itemFoundBack != null)
                {
                    // Found item in both
                    // Recursively Enumerate all!
                    GetElements(itemElement.FindAllBuildCache(TreeScope.TreeScope_Children,
                            _automation.CreateTrueCondition(), _cacheRequest))
                        .Apply(curItem => addContainerItems(curItem, true));
                }
                else if (itemFoundFront != null)
                {
                    // Found item in forward phase
                    // Recursively Enumerate forward
                    GetNextSiblingElementsBuildCache(itemFoundFront)
                        .TakeWhile(e => e.CachedIsOffscreen == 0)
                        .Apply(curItem => addContainerItems(curItem, true));
                }
                else if (!itemLevel && itemFoundBack != null)
                {
                    // Found item in backward phase
                    // Recursively Enumerate backward
                    IUIAutomationElement frontElement = null;
                    GetPreviousSiblingElementsBuildCache(itemFoundBack)
                        .TakeWhile(
                            e =>
                            {
                                if (e.CachedIsOffscreen == 0)
                                {
                                    return true;
                                }
                                frontElement = e;
                                return false;
                            })
                        .Apply(curElement => addContainerItems(curElement, true));
                    
                    // descendant of pre-item1 may be on-screen
                    if (frontElement != null)
                    {
                        for (var cur2 = frontElement;
                            cur2 != null;
                            cur2 = _treeWalker.GetLastChildElementBuildCache(cur2, _cacheRequest))
                        {
                            if (cur2.CachedIsOffscreen == 0)
                            {
                                // Enumerate backward from item2
                                GetPreviousSiblingElementsBuildCache(cur2)
                                    .TakeWhile(
                                        e => e.CachedIsOffscreen == 0)
                                    .Apply(curElement => addContainerItems(curElement, true));
                                break;
                            }
                        }
                    }
                }
                else if (!itemLevel)
                {
                    // Search screen diagonally to find an on-screen item, e.g. item3
                    Func<tagPOINT, List<IUIAutomationElement>> getItemAncestorsFromPoint = point =>
                    {
                        var elem = _automation.ElementFromPointBuildCache(point, _cacheRequest);
                        if (elem == null)
                        {
                            return null;
                        }

                        var itemAncestors = new List<IUIAutomationElement>();
                        var isItemDescendant = false;

                        for (var curElement = elem;
                            curElement != null;
                            curElement = _treeWalker.GetParentElementBuildCache(curElement, _cacheRequest))
                        {
                            if (_automation.CompareElements(curElement, itemElement) != 0)
                            {
                                isItemDescendant = true;
                                break;
                            }
                            else if (isItemElement(curElement))
                            {
                                itemAncestors.Add(curElement);
                            }
                        }

                        if (isItemDescendant && itemAncestors.Count > 0)
                        {
                            itemAncestors.Reverse();
                            return itemAncestors;
                        }

                        return null;
                    };

                    List<IUIAutomationElement> foundElementAncestors = null;
                    var boundRect = itemElement.CurrentBoundingRectangle;
                    var delta = 8.0 / (boundRect.right - boundRect.left);

                    for (double t = 0; t < 1.0; t += delta)
                    {
                        tagPOINT p;
                        p.x = (int)Math.Floor(
                            boundRect.left + t * (boundRect.right - boundRect.left));
                        p.y = (int)Math.Floor(
                            boundRect.top + t * (boundRect.bottom - boundRect.top));

                        foundElementAncestors = getItemAncestorsFromPoint(p);
                        if (foundElementAncestors != null)
                        {
                            break;
                        }
                    }

                    if (foundElementAncestors == null)
                    {
                        // Still found none
                        // Done!
                        return;
                    }

                    var curLevel = 0;
                    foreach (var cur in foundElementAncestors)
                    {
                        ++curLevel;
                        if (cur.CachedIsOffscreen == 0)
                        {
                            // If item1 is on-screen
                            // Enumerate backward/forward from item1
                            IUIAutomationElement frontElement = null;
                            GetPreviousSiblingElementsBuildCache(
                                    _treeWalker.GetPreviousSiblingElementBuildCache(cur, _cacheRequest))
                                .TakeWhile(
                                    e =>
                                    {
                                        if (e.CachedIsOffscreen == 0)
                                        {
                                            return true;
                                        }
                                        frontElement = e;
                                        return false;
                                    })
                                .Apply(curElement => addContainerItems(curElement, true)); // Enumerate from level2 (forward+backward, forward)

                            // descendant of pre-item1 may be on-screen
                            // pre-foundElementAncestors[0] if item1 is off-screen
                            // pre-frontElement if item1 is on-screen
                            if (curLevel == 1 && frontElement != null || curLevel > 1)
                            {
                                var preItem = curLevel == 1
                                    ? frontElement
                                    : _treeWalker.GetPreviousSiblingElementBuildCache(foundElementAncestors[0], _cacheRequest);
                                // If previous item1 is off-screen
                                //  If last item2 is off-screen
                                //          ...
                                for (var cur2 = preItem;
                                    cur2 != null;
                                    cur2 = _treeWalker.GetLastChildElementBuildCache(cur2, _cacheRequest))
                                {
                                    if (cur2.CachedIsOffscreen == 0)
                                    {
                                        //      Enumerate backward from item2
                                        GetPreviousSiblingElementsBuildCache(cur2)
                                            .TakeWhile(e => e.CachedIsOffscreen == 0)
                                            .Apply(curElement => addContainerItems(curElement, true));
                                        break;
                                    }
                                }
                            }

                            GetNextSiblingElementsBuildCache(cur)
                                .TakeWhile(e => e.CachedIsOffscreen == 0)
                                .Apply(curElement => addContainerItems(curElement, true));
                            break;
                        }
                        else
                        {
                            // descendant of post-item1 may be on-screen
                            GetNextSiblingElementsBuildCache(
                                    _treeWalker.GetNextSiblingElementBuildCache(cur, _cacheRequest))
                                .TakeWhile(e => e.CachedIsOffscreen == 0)
                                .Apply(curElement => addContainerItems(curElement, true));
                        }
                    }
                }
            };

            addElements(_automation.ElementFromHandleBuildCache(hWnd, _cacheRequest));
            return result;
        }

        private static IEnumerable<IUIAutomationElement> GetElements(IUIAutomationElementArray elements)
        {
            for (var i = 0; i < elements.Length; ++i)
            {
                yield return elements.GetElement(i);
            }
        }

        private IEnumerable<IUIAutomationElement> GetPreviousSiblingElementsBuildCache(IUIAutomationElement element)
        {
            for (var curElement = element;
                curElement != null;
                curElement = _treeWalker.GetPreviousSiblingElementBuildCache(curElement, _cacheRequest))
            {
                yield return curElement;
            }
        }

        private IEnumerable<IUIAutomationElement> GetNextSiblingElementsBuildCache(IUIAutomationElement element)
        {
            for (var curElement = element;
                curElement != null;
                curElement = _treeWalker.GetNextSiblingElementBuildCache(curElement, _cacheRequest))
            {
                yield return curElement;
            }
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
    }
}
