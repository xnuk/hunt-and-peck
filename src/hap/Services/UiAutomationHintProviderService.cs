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
        // TODO: wrap in method
        private static readonly int[] s_containerPropertyIds =
        {
            UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsSpreadsheetPatternAvailablePropertyId,
            UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId
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

        private readonly IUIAutomation _automation;

        private readonly IUIAutomationCondition _conditionAccessibleControl;

        private readonly IUIAutomationCondition _conditionContainer;

        private readonly IUIAutomationTreeWalker _controlViewWalker;

        private readonly IUIAutomationTreeWalker _itemTreeWalker;

        private readonly IUIAutomationCacheRequest _itemCacheRequest;

        private static readonly int[] s_noFocusControlTypeIds =
        {
            UIA_ControlTypeIds.UIA_GroupControlTypeId,
            UIA_ControlTypeIds.UIA_PaneControlTypeId,
            UIA_ControlTypeIds.UIA_WindowControlTypeId
        };

        public UiAutomationHintProviderService()
        {
            _automation = new CUIAutomation();
            
            var conditionEnabledControl = _automation.CreateAndCondition(
                _automation.ControlViewCondition,
                _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsEnabledPropertyId, true));

            _conditionAccessibleControl = _automation.CreateAndCondition(
                conditionEnabledControl,
                _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsOffscreenPropertyId, false)
            );

            _conditionContainer = _automation.CreateAndConditionFromArray(new[]
            {
                // A large container should support scroll pattern.
                _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId, true),
                // Exclude by control type
                _automation.CreateNotCondition(_automation.CreateOrConditionFromArray(new[]
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
                }.Select(id => _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, id)).ToArray())),
                // Include by pattern availability
                _automation.CreateOrConditionFromArray(new[]
                {
                    UIA_PropertyIds.UIA_IsGridPatternAvailablePropertyId,
                    UIA_PropertyIds.UIA_IsItemContainerPatternAvailablePropertyId,
                    UIA_PropertyIds.UIA_IsScrollPatternAvailablePropertyId,
                    UIA_PropertyIds.UIA_IsSpreadsheetPatternAvailablePropertyId,
                    UIA_PropertyIds.UIA_IsTablePatternAvailablePropertyId
                }.Select(id => _automation.CreatePropertyCondition(id, true)).ToArray())
            });

            _controlViewWalker = _automation.ControlViewWalker;

            _itemTreeWalker = _automation.CreateTreeWalker(conditionEnabledControl);

            _itemCacheRequest = _automation.CreateCacheRequest();
            _itemCacheRequest.AddProperty(UIA_PropertyIds.UIA_IsOffscreenPropertyId);
            foreach (var propertyId in s_containerItemPropertyIds)
            {
                _itemCacheRequest.AddProperty(propertyId);
            }
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
                var boundingRectObject = element.CurrentBoundingRectangle;
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
            var automationElement = _automation.ElementFromHandle(hWnd);

            //var conditionControlView = _automation.ControlViewCondition;
            //var conditionEnabled = _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsEnabledPropertyId, true);
            //var enabledControlCondition = _automation.CreateAndCondition(conditionControlView, conditionEnabled);

            //var conditionOnScreen = _automation.CreatePropertyCondition(UIA_PropertyIds.UIA_IsOffscreenPropertyId, false);
            //var condition = _automation.CreateAndCondition(enabledControlCondition, conditionOnScreen);

            //var elementArray = automationElement.FindAll(TreeScope.TreeScope_Descendants, condition);
            //for (var i = 0; i < elementArray.Length; ++i)
            //{
            //    result.Add(elementArray.GetElement(i));
            //}
            AddElements(automationElement, result);

            return result;
        }

        /// <summary>
        /// Adds automation elements in the subtree of <paramref name="automationElement"/> to <paramref name="result"/>.
        /// </summary>
        private void AddElements(IUIAutomationElement automationElement, List<IUIAutomationElement> result)
        {
            var element = automationElement;

            var container = element.FindFirst(TreeScope.TreeScope_Descendants, _conditionContainer);
            if (container == null)
            {
                result.AddRange(element.FindAll(
                    TreeScope.TreeScope_Subtree, _conditionAccessibleControl).AsEnumerable());
                return;
            }

            var containerAncestors = new List<IUIAutomationElement>();
            for (var curParent = container;
                _automation.CompareElements(curParent, element) == 0;
                curParent = _controlViewWalker.GetParentElement(curParent))
            {
                containerAncestors.Add(curParent);
            }
            containerAncestors.Add(element);
            containerAncestors.Reverse();

            foreach (var l in Enumerable.Range(0, containerAncestors.Count - 1))
            {
                var nextElem = _controlViewWalker.EnumerateChildren(containerAncestors[l])
                    .SkipWhile(curElem =>
                    {
                        if (_automation.CompareElements(curElem, containerAncestors[l + 1]) == 0)
                        {
                            result.AddRange(curElem.FindAll(
                                TreeScope.TreeScope_Subtree, _conditionAccessibleControl).AsEnumerable());
                            return true;
                        }

                        if (l + 1 == containerAncestors.Count - 1)
                        {
                            AddContainerElements(containerAncestors[l + 1], result);
                        }
                        else
                        {
                            result.AddRange(curElem.FindAll(
                                TreeScope.TreeScope_Element, _conditionAccessibleControl).AsEnumerable());
                        }
                        return false;
                    }).Skip(1).FirstOrDefault();

                if (nextElem == null)
                    continue;

                foreach (var child in _controlViewWalker.EnumerateSiblingsFrom(nextElem))
                {
                    AddElements(child, result);
                }
            }
        }

        /// <summary>
        /// Adds automation elements in the subtree of <paramref name="containerElement"/> to <paramref name="result"/>.
        /// </summary>
        private void AddContainerElements(IUIAutomationElement containerElement, List<IUIAutomationElement> result)
        {
            var accessibleControl = containerElement.FindAll(TreeScope.TreeScope_Element, _conditionAccessibleControl)
                .AsEnumerable().FirstOrDefault();
            if (accessibleControl == null)
            {
                return;
            }

            var scrollPattern = (IUIAutomationScrollPattern) containerElement.GetCurrentPattern(
                UIA_PatternIds.UIA_ScrollPatternId);

            if (scrollPattern.CurrentHorizontalViewSize * scrollPattern.CurrentVerticalViewSize / 10000.0 >= 0.8)
            {
                result.AddRange(containerElement.FindAll(TreeScope.TreeScope_Subtree,
                    _conditionAccessibleControl).AsEnumerable());
                return;
            }

            var itemFound = AddContainerSubItems(containerElement, result);
            List<IUIAutomationElement> itemFoundAncestors = null;

            if (itemFound == containerElement)
            {
                // Found all children
                return;
            }
            else if (itemFound != null)
            {
                // Children at front may be off-screen, check from back
                itemFoundAncestors = new List<IUIAutomationElement> {itemFound};
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
                    point.x = (int) Math.Floor(
                        boundRect.left + t * (boundRect.right - boundRect.left));
                    point.y = (int) Math.Floor(
                        boundRect.top + t * (boundRect.bottom - boundRect.top));

                    var foundElement = _automation.ElementFromPointBuildCache(point, _itemCacheRequest);
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
                        curParent = _itemTreeWalker.GetParentElementBuildCache(curParent, _itemCacheRequest))
                    {
                        if (_automation.CompareElements(curParent, containerElement) != 0)
                        {
                            isContainerDescendant = true;
                            break;
                        }
                        else if (CachedIsItemElement(curParent))
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
                    var itemFirstOffscreen = _itemTreeWalker.EnumerateSiblingsFromBackwardBuildCache(itemAncestor, _itemCacheRequest)
                        .Skip(1)
                        .FirstOrDefault(
                            item =>
                            {
                                if (item.CachedIsOffscreen == 0)
                                {
                                    AddContainerSubItems(item, result);
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
                            : _itemTreeWalker.GetPreviousSiblingElementBuildCache(itemFoundAncestors[0], _itemCacheRequest);
                        // One of last childs should be on-screen
                        for (var curChild = itemOffscreen;
                            curChild != null;
                            curChild = _itemTreeWalker.GetLastChildElementBuildCache(curChild, _itemCacheRequest))
                        {
                            if (curChild.CachedIsOffscreen == 0)
                            {
                                // Find item backward from curChild
                                foreach (var item in _itemTreeWalker.EnumerateSiblingsFromBackwardBuildCache(
                                    curChild, _itemCacheRequest))
                                {
                                    if (item.CachedIsOffscreen != 0)
                                    {
                                        break;
                                    }
                                    AddContainerSubItems(item, result);
                                }
                                
                                break;
                            }
                        }
                    }

                    // Find item forward from itemAncestor
                    foreach (var item in _itemTreeWalker.EnumerateSiblingsFromBuildCache(itemAncestor, _itemCacheRequest))
                    {
                        if (item.CachedIsOffscreen != 0)
                        {
                            break;
                        }
                        AddContainerSubItems(item, result);
                    }

                    break;
                }
                else
                {
                    // Siblings next to itemAncestor may be on-screen
                    foreach (var item in _itemTreeWalker.EnumerateSiblingsFromBuildCache(itemAncestor, _itemCacheRequest)
                        .Skip(1))
                    {
                        if (item.CachedIsOffscreen != 0)
                        {
                            break;
                        }
                        AddContainerSubItems(item, result);
                    }
                }
            }
        }

        /// <summary>
        /// Adds container items in the subtree of <paramref name="itemElement"/> to <paramref name="result"/>.
        /// </summary>
        /// <returns></returns>
        private IUIAutomationElement AddContainerSubItems(IUIAutomationElement itemElement,
            List<IUIAutomationElement> result)
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
                        if (CachedIsItemElement(element))
                        {
                            // Found an on-screen item
                            return true;
                        }

                        // Find non-item elements
                        AddElements(element, result);
                        return false;
                    });

            // Find an on-screen item from front
            var itemFoundFront = findFirstItemAndNonItems(
                _itemTreeWalker.EnumerateChildrenBuildCache(itemElement, _itemCacheRequest));

            // Find an on-screen item from back
            var itemFoundBack = findFirstItemAndNonItems(
                _itemTreeWalker.EnumerateChildrenBackwardBuildCache(itemElement, _itemCacheRequest));

            if (itemFoundFront != null && itemFoundBack != null)
            {
                // Found item from both, add all children
                foreach (var item in itemElement.FindAll(TreeScope.TreeScope_Children, _conditionAccessibleControl)
                    .AsEnumerable())
                {
                    AddContainerSubItems(item, result);
                }
                return itemElement;
            }
            else if (itemFoundFront != null)
            {
                // Add children from front
                foreach (var item in _itemTreeWalker.EnumerateSiblingsFrom(itemFoundFront))
                {
                    if (item.CachedIsOffscreen != 0)
                    {
                        break;
                    }
                    AddContainerSubItems(item, result);
                }
                return itemElement;
            }

            return itemFoundBack;
        }

        /// <summary>
        /// Decides whether <paramref name="element"/> is an item in a container using the cache.
        /// </summary>
        private bool CachedIsItemElement(IUIAutomationElement element)
        {
            return s_containerItemPropertyIds.Any(propertyId => (bool) element.GetCachedPattern(propertyId));
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
                var invokePattern = (IUIAutomationInvokePattern)automationElement.GetCurrentPattern(
                    UIA_PatternIds.UIA_InvokePatternId);
                if (invokePattern != null)
                {
                    return new UiAutomationInvokeHint(owningWindow, invokePattern, hintBounds);
                }

                var togglePattern = (IUIAutomationTogglePattern)automationElement.GetCurrentPattern(
                    UIA_PatternIds.UIA_TogglePatternId);
                if (togglePattern != null)
                {
                    return new UiAutomationToggleHint(owningWindow, togglePattern, hintBounds);
                }

                var isFocusable = automationElement.CurrentIsKeyboardFocusable != 0;
                if (isFocusable && !s_noFocusControlTypeIds.Contains(automationElement.CurrentControlType))
                {
                    return new UiAutomationFocusHint(owningWindow, automationElement, hintBounds);
                }
                // TODO: noncontainer
                // TODO: focus part of item?
                var isContainer = s_containerPropertyIds.Any(
                    propertyId => (bool)automationElement.GetCurrentPropertyValue(propertyId));
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
