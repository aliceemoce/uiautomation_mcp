# -*- coding: utf-8 -*-
"""Interaction operation tools."""

import logging
from typing import Optional

from mcp.server.fastmcp import FastMCP

import uiautomation as auto

from ..core import (
    get_control_by_handle,
    format_error,
    create_confirmation,
    confirm_operation,
    check_admin,
)
from ..config import config

logger = logging.getLogger(__name__)

# Store pending confirmations for dangerous operations
_pending_confirms = {}

# uiautomation's Click() always calls SetCursorPos internally.
# Work around by saving/restoring cursor position around every click.
def _safe_click(ctrl, func, ratioX=0.5, ratioY=0.5):
    import win32api
    try:
        old = win32api.GetCursorPos()
    except Exception:
        old = None
    func(ctrl, ratioX=ratioX, ratioY=ratioY, simulateMove=False, waitTime=0)
    if old is not None:
        try:
            win32api.SetCursorPos(*old)
        except Exception:
            pass


def register_interaction_tools(mcp: FastMCP):
    """Register interaction tools with the MCP server."""

    @mcp.tool()
    def ui_click(
        handle: Optional[int] = None,
        x: Optional[int] = None,
        y: Optional[int] = None,
        button: str = "left",
        double: bool = False,
    ) -> dict:
        """Click on a control or at coordinates.

        NO physical cursor movement - uses UIA ratio-click (simulateMove=False).

        Args:
            handle: Control handle to click (uses center if no x/y)
            x: Relative X offset from control top-left, or absolute screen X
            y: Relative Y offset from control top-left, or absolute screen Y
            button: Mouse button (left, right, middle)
            double: Whether to double-click

        Returns:
            Success or error
        """
        check_admin()

        try:
            # Click on control (with optional x/y pixel offset)
            if handle is not None:
                control = get_control_by_handle(handle)
                if not control:
                    return format_error(
                        "CONTROL_NOT_FOUND",
                        f"Invalid control handle: {handle}",
                    )

                if x is not None and y is not None:
                    try:
                        rect = control.BoundingRectangle
                        cw = max(1, rect.width())
                        ch = max(1, rect.height())
                        rx = max(0.01, min(0.99, x / cw))
                        ry = max(0.01, min(0.99, y / ch))
                    except Exception:
                        rx, ry = 0.5, 0.5
                else:
                    rx, ry = 0.5, 0.5

                if button == "right":
                    _safe_click(control, lambda c, **kw: c.RightClick(**kw), ratioX=rx, ratioY=ry)
                elif button == "middle":
                    _safe_click(control, lambda c, **kw: c.MiddleClick(**kw), ratioX=rx, ratioY=ry)
                elif double:
                    _safe_click(control, lambda c, **kw: c.DoubleClick(**kw), ratioX=rx, ratioY=ry)
                else:
                    _safe_click(control, lambda c, **kw: c.Click(**kw), ratioX=rx, ratioY=ry)

                return {"success": True, "data": {"action": "click", "handle": handle, "ratioX": rx, "ratioY": ry}}

            # Absolute-coordinate click: resolve via UIA ControlFromPoint, then ratio-click
            if x is not None and y is not None:
                ctrl = auto.ControlFromPoint(x, y)
                if ctrl is None:
                    return format_error(
                        "NO_CONTROL_AT_POINT",
                        f"No control at ({x}, {y})",
                        ["Try providing a handle instead"],
                    )
                try:
                    rect = ctrl.BoundingRectangle
                    cw = max(1, rect.width())
                    ch = max(1, rect.height())
                    rx = max(0.01, min(0.99, (x - int(rect.left)) / cw))
                    ry = max(0.01, min(0.99, (y - int(rect.top)) / ch))
                except Exception:
                    rx, ry = 0.5, 0.5

                if button == "right":
                    _safe_click(ctrl, lambda c, **kw: c.RightClick(**kw), ratioX=rx, ratioY=ry)
                elif button == "middle":
                    _safe_click(ctrl, lambda c, **kw: c.MiddleClick(**kw), ratioX=rx, ratioY=ry)
                elif double:
                    _safe_click(ctrl, lambda c, **kw: c.DoubleClick(**kw), ratioX=rx, ratioY=ry)
                else:
                    _safe_click(ctrl, lambda c, **kw: c.Click(**kw), ratioX=rx, ratioY=ry)

                return {"success": True, "data": {"action": "click", "x": x, "y": y, "controlType": ctrl.ControlTypeName}}

            return format_error(
                "INVALID_PARAMS",
                "Need handle or (x, y) coordinates",
            )

        except Exception as e:
            logger.exception("ui_click failed")
            return format_error("INTERNAL_ERROR", str(e))

    @mcp.tool()
    def ui_send_keys(
        handle: int,
        text: str,
        interval: float = 0.05,
    ) -> dict:
        """Send keyboard input to a control.

        Args:
            handle: Control handle
            text: Text/keys to send (use {Ctrl}, {Enter}, etc. for special keys)
            interval: Interval between keystrokes in seconds

        Returns:
            Success or error
        """
        check_admin()

        try:
            control = get_control_by_handle(handle)
            if not control:
                return format_error(
                    "CONTROL_NOT_FOUND",
                    f"Invalid control handle: {handle}",
                )

            control.SendKeys(text, interval=interval)
            return {"success": True, "data": {"action": "send_keys", "text": text}}

        except Exception as e:
            logger.exception("ui_send_keys failed")
            return format_error("INTERNAL_ERROR", str(e))

    @mcp.tool()
    def ui_set_value(
        handle: int,
        value: str,
    ) -> dict:
        """Set text value of a control using ValuePattern.

        Args:
            handle: Control handle
            value: Value to set

        Returns:
            Success or error
        """
        check_admin()

        try:
            control = get_control_by_handle(handle)
            if not control:
                return format_error(
                    "CONTROL_NOT_FOUND",
                    f"Invalid control handle: {handle}",
                )

            pattern = control.GetValuePattern()
            if not pattern:
                return format_error(
                    "PATTERN_NOT_SUPPORTED",
                    "Control does not support ValuePattern",
                    ["Try ui_send_keys instead"],
                )

            pattern.SetValue(value)
            return {"success": True, "data": {"action": "set_value", "value": value}}

        except Exception as e:
            logger.exception("ui_set_value failed")
            return format_error("INTERNAL_ERROR", str(e))

    @mcp.tool()
    def ui_close_window(
        handle: int,
        confirmationToken: Optional[str] = None,
    ) -> dict:
        """Close a window. Requires confirmation.

        Args:
            handle: Window handle
            confirmationToken: Token from previous confirmation (if required)

        Returns:
            Confirmation request, success, or error
        """
        check_admin()

        try:
            control = get_control_by_handle(handle)
            if not control:
                return format_error(
                    "CONTROL_NOT_FOUND",
                    f"Invalid window handle: {handle}",
                )

            # Check if confirmation is needed
            if config.confirmation_enabled and not confirmationToken:
                request = create_confirmation(
                    "ui_close_window",
                    {"windowName": control.Name, "handle": handle},
                    f"About to close window '{control.Name}', continue?",
                )
                return {"success": False, "requiresConfirmation": True, "confirmation": request.model_dump()}

            # Verify confirmation token
            if config.confirmation_enabled and confirmationToken:
                result = confirm_operation(confirmationToken, True)
                if not result:
                    return format_error("INVALID_CONFIRMATION", "Confirmation token invalid or expired")

            # Close the window
            pattern = control.GetWindowPattern()
            if pattern:
                pattern.Close()
            else:
                # Fallback to Alt+F4
                control.SetFocus()
                auto.SendKeys("{Alt}{F4}")

            return {"success": True, "data": {"action": "close_window", "handle": handle}}

        except Exception as e:
            logger.exception("ui_close_window failed")
            return format_error("INTERNAL_ERROR", str(e))

    @mcp.tool()
    def ui_move_window(
        handle: int,
        x: Optional[int] = None,
        y: Optional[int] = None,
        width: Optional[int] = None,
        height: Optional[int] = None,
    ) -> dict:
        """Move and/or resize a window.

        Args:
            handle: Window handle
            x: New X position (optional)
            y: New Y position (optional)
            width: New width (optional)
            height: New height (optional)

        Returns:
            Success or error
        """
        check_admin()

        try:
            control = get_control_by_handle(handle)
            if not control:
                return format_error(
                    "CONTROL_NOT_FOUND",
                    f"Invalid window handle: {handle}",
                )

            control.MoveWindow(x, y, width, height)
            return {
                "success": True,
                "data": {
                    "action": "move_window",
                    "x": x,
                    "y": y,
                    "width": width,
                    "height": height,
                }
            }

        except Exception as e:
            logger.exception("ui_move_window failed")
            return format_error("INTERNAL_ERROR", str(e))
