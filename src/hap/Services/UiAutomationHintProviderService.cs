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
        private static readonly int[] s_containerPropertyIds =
        {
            UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSpreadsheetPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId
        };

        private static readonly int[] s_nonContainerControlTypeIds =
        {
            UIA_ControlTypeIds.UIA_ComboBoxControlTypeId,
            UIA_ControlTypeIds.UIA_CalendarControlTypeId,
            UIA_ControlTypeIds.UIA_DocumentControlTypeId,
            UIA_ControlTypeIds.UIA_EditControlTypeId,
            UIA_ControlTypeIds.UIA_GroupControlTypeId,
            UIA_ControlTypeIds.UIA_MenuControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_TabControlTypeId,
            UIA_ControlTypeIds.UIA_SliderControlTypeId,
            UIA_ControlTypeIds.UIA_SpinnerControlTypeId,
            UIA_ControlTypeIds.UIA_SplitButtonControlTypeId
        };

        private static readonly int[] s_containerItemPropertyIds =
        {
            UIA_PropertyIds.UIA_IsExpandCollapsePatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsGridItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsScrollItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSelectionItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSpreadsheetItemPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsTableItemPatternAvailablePropertyId
        };

        private static readonly int[] s_noFocusControlTypeIds =
        {
            UIA_ControlTypeIds.UIA_GroupControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_WindowControlTypeId
        };

        private readonly IUIAutomation _automation;

        private readonly IUIAutomationTreeWalker _treeWalker;

        private readonly IUIAutomationCondition _automationCondition;

        private readonly IUIAutomationCacheRequest _cacheRequest;

        public UiAutomationHintProviderService()
        {
            _automation = new CUIAutomation();

            var baseCondition = _automation.CreateAndConditionFromArray(new[]
            {
                _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsControlElementPropertyId, true),
                _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsEnabledPropertyId, true)
            });

            _treeWalker = _automation.CreateTreeWalker(baseCondition);

            _automationCondition = _automation.CreateAndCondition(baseCondition,
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
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_IsOffscreenPropertyId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_ScrollHorizontalViewSizePropertyId);
            _cacheRequest.AddProperty(UIA_PropertyIds.UIA_ScrollVerticalViewSizePropertyId);
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

            // TODO: wrap in method
            Func<IUIAutomationElement, bool> isItemElement = element => s_containerItemPropertyIds.Any(
                propertyId => (bool)element.GetCachedPropertyValue(propertyId));

            // Assume itemElement is on-screen,
            // Add its on-screen subitems including itself.
            Func<IUIAutomationElement, IUIAutomationElement> addSubItems = null;
            addSubItems = itemElement =>
            {
                result.Add(itemElement);

                // Assume structure
                //  (Header)
                //  item(s)
                //  (Footer)
                var elementIds = new HashSet<string>();
                Func<IUIAutomationElement, bool> addNewElement = element => elementIds.Add(
                    string.Join(".", from subId in (int[])element.GetRuntimeId() select subId.ToString()));
                Func<IEnumerable<IUIAutomationElement>, IUIAutomationElement> findFirstItemAndNonItems = elements =>
                    elements.TakeWhile(element => element.CachedIsOffscreen == 0 && addNewElement(element))
                            .FirstOrDefault(element =>
                            {
                                if (isItemElement(element))
                                {
                                    // Found an on-screen item
                                    return true;
                                }
                                // Find non-item elements
                                addElements(element);
                                return false;
                            });

                // Find an on-screen item from front
                var itemFoundFront = findFirstItemAndNonItems(
                    EnumNextSiblingElementsBuildCache(_treeWalker.GetFirstChildElementBuildCache(
                        itemElement, _cacheRequest)));

                // Find an on-screen item from back
                var itemFoundBack = findFirstItemAndNonItems(
                    EnumPreviousSiblingElementsBuildCache(_treeWalker.GetLastChildElementBuildCache(
                        itemElement, _cacheRequest)));

                if (itemFoundFront != null && itemFoundBack != null)
                {
                    // Found item from both, add all children
                    EnumElementArray(itemElement.FindAllBuildCache(TreeScope.TreeScope_Children,
                            _treeWalker.condition, _cacheRequest))
                        .Apply(item => addSubItems(item));

                    return itemElement;
                }
                else if (itemFoundFront != null)
                {
                    // Add children from front
                    EnumNextSiblingElementsBuildCache(itemFoundFront)
                        .TakeWhile(element => element.CachedIsOffscreen == 0)
                        .Apply(item => addSubItems(item));

                    return itemElement;
                }

                return itemFoundBack;
            };

            // Add on-screen items of a typical list or tree, including itself.
            Action<IUIAutomationElement> addContainerItems = containerElement =>
            {
                var itemFound = addSubItems(containerElement);
                List<IUIAutomationElement> itemFoundAncestors = null;

                if (itemFound == containerElement)
                {
                    // Found all children
                    return;
                }
                else if (itemFound != null)
                {
                    // Children at front may be off-screen, check from back
                    itemFoundAncestors = new List<IUIAutomationElement> { itemFound };
                }
                else
                {
                    // Scan screen diagonally to find an on-screen item
                    var boundRect = containerElement.CurrentBoundingRectangle;
                    // TODO: adjust
                    var delta = 8.0 / (boundRect.right - boundRect.left);

                    for (double t = 0; t < 1.0; t += delta)
                    {
                        tagPOINT point;
                        point.x = (int)Math.Floor(
                            boundRect.left + t * (boundRect.right - boundRect.left));
                        point.y = (int)Math.Floor(
                            boundRect.top + t * (boundRect.bottom - boundRect.top));

                        var foundElement = _automation.ElementFromPointBuildCache(point, _cacheRequest);
                        if (foundElement == null)
                        {
                            continue;
                        }

                        // Check foundElement is a descendant of container
                        // Retrieve ancestor items
                        var isContainerDescendant = false;
                        var itemAncestors = new List<IUIAutomationElement>();

                        for (var curParent = foundElement;
                            curParent != null;
                            curParent = _treeWalker.GetParentElementBuildCache(curParent, _cacheRequest))
                        {
                            if (_automation.CompareElements(curParent, containerElement) != 0)
                            {
                                isContainerDescendant = true;
                                break;
                            }
                            else if (isItemElement(curParent))
                            {
                                itemAncestors.Add(curParent);
                            }
                        }

                        if (isContainerDescendant && itemAncestors.Count > 0)
                        {
                            itemAncestors.Reverse();
                            itemFoundAncestors = itemAncestors;
                            break;
                        }
                    }
                }

                if (itemFoundAncestors == null)
                {
                    // Give up
                    return;
                }

                var itemLevel = 0;
                foreach (var itemAncestor in itemFoundAncestors)
                {
                    ++itemLevel;
                    if (itemAncestor.CachedIsOffscreen == 0)
                    {
                        // Find item backward from itemAncestor
                        var itemFirstOffscreen = EnumPreviousSiblingElementsBuildCache(
                                _treeWalker.GetPreviousSiblingElementBuildCache(itemAncestor, _cacheRequest))
                            .FirstOrDefault(
                                item =>
                                {
                                    if (item.CachedIsOffscreen == 0)
                                    {
                                        addSubItems(item);
                                        return false;
                                    }
                                    return true;
                                });

                        // Descendant of...
                        //  itemFirstOffscreen if foundItemAncestors[0] is on-screen
                        //  previous sibling of foundItemAncestors[0] otherwise
                        // ...may be on-screen
                        if (itemLevel == 1 && itemFirstOffscreen != null || itemLevel > 1)
                        {
                            var itemOffscreen = itemLevel == 1
                                ? itemFirstOffscreen
                                : _treeWalker.GetPreviousSiblingElementBuildCache(itemFoundAncestors[0], _cacheRequest);
                            // One of last childs should be on-screen
                            for (var curChild = itemOffscreen;
                                curChild != null;
                                curChild = _treeWalker.GetLastChildElementBuildCache(curChild, _cacheRequest))
                            {
                                if (curChild.CachedIsOffscreen == 0)
                                {
                                    // Find item backward from curChild
                                    EnumPreviousSiblingElementsBuildCache(curChild)
                                        .TakeWhile(item => item.CachedIsOffscreen == 0)
                                        .Apply(item => addSubItems(item));
                                    break;
                                }
                            }
                        }

                        // Find item forward from itemAncestor
                        EnumNextSiblingElementsBuildCache(itemAncestor)
                            .TakeWhile(item => item.CachedIsOffscreen == 0)
                            .Apply(item => addSubItems(item));
                        break;
                    }
                    else
                    {
                        // Siblings next to itemAncestor may be on-screen
                        EnumNextSiblingElementsBuildCache(
                                _treeWalker.GetNextSiblingElementBuildCache(itemAncestor, _cacheRequest))
                            .TakeWhile(item => item.CachedIsOffscreen == 0)
                            .Apply(item => addSubItems(item));
                    }
                }
            };

            var level = 0;
            addElements = element =>
            {
                ++level;
                var sw = new Stopwatch();
                sw.Start();

                Func<IUIAutomationElement, bool> isLargeContainer = element2 =>
                {
                    if (s_nonContainerControlTypeIds.Contains(element2.CachedControlType))
                    {
                        return false;
                    }

                    if (s_containerPropertyIds.All(propertyId => !(bool) element2.GetCachedPropertyValue(propertyId)))
                    {
                        return false;
                    }

                    if (!(bool)element2.GetCachedPropertyValue(UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId))
                    {
                        return false;
                    }

                    return (double)element2.GetCachedPropertyValue(UIA_PropertyIds.UIA_ScrollHorizontalViewSizePropertyId)
                        * (double)element2.GetCachedPropertyValue(UIA_PropertyIds.UIA_ScrollVerticalViewSizePropertyId)
                        / 10000.0 < 0.8;
                };

                var lc = isLargeContainer(element);

                if (!lc)
                {
                    result.Add(element);
                    EnumElementArray(element.FindAllBuildCache(TreeScope.TreeScope_Children, _automationCondition, _cacheRequest))
                        .Apply(child => addElements(child));
                }
                else
                {
                    // A container may have a large number of off-screen elements.
                    addContainerItems(element);
                }

                --level;
                sw.Stop();
                Debug.Write(new string(' ', level*2));
                if (lc)
                    Debug.Write("Z");
                Debug.WriteLine("{0} : {1} took {2} ms", string.Join(".", from subId in (int[])element.GetRuntimeId() select subId.ToString("X")), element.CurrentName, sw.ElapsedMilliseconds);
            };

            addElements(_automation.ElementFromHandleBuildCache(hWnd, _cacheRequest));
            return result;
        }

        /// <summary>
        /// Enumerates the array of elements
        /// </summary>
        private static IEnumerable<IUIAutomationElement> EnumElementArray(IUIAutomationElementArray elements)
        {
            for (var i = 0; i < elements.Length; ++i)
            {
                yield return elements.GetElement(i);
            }
        }

        /// <summary>
        /// Enumerates siblings previous to the element, including itself
        /// </summary>
        private IEnumerable<IUIAutomationElement> EnumPreviousSiblingElementsBuildCache(IUIAutomationElement element)
        {
            for (var curElement = element;
                curElement != null;
                curElement = _treeWalker.GetPreviousSiblingElementBuildCache(curElement, _cacheRequest))
            {
                yield return curElement;
            }
        }

        /// <summary>
        /// Enumerates siblings next to the element, including itself
        /// </summary>
        private IEnumerable<IUIAutomationElement> EnumNextSiblingElementsBuildCache(IUIAutomationElement element)
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
                var invokePattern = (IUIAutomationInvokePattern)automationElement.GetCachedPattern(
                    UIA_PatternIds.UIA_InvokePatternId);
                if (invokePattern != null)
                {
                    return new UiAutomationInvokeHint(owningWindow, invokePattern, hintBounds);
                }

                var togglePattern = (IUIAutomationTogglePattern)automationElement.GetCachedPattern(
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
                    propertyId => (bool)automationElement.GetCachedPropertyValue(propertyId));
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
                    if (pattern != null)
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
