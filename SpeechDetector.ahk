﻿; CODE ADAPTED FROM https://www.autohotkey.com/board/topic/38015-ahkhid-an-ahk-implementation-of-the-hid-functions/
; Note: for this script to work, it must be one of the first to be included in the main script!
; Be carefull with labels in the main script (or includes other scripts), as this can cause interference and may casue this script to fail to work.

;Must be in auto-execute section
#Include "AHKHID.ahk"

; PHILIPS SPEECH DEVICE IDENTIFICATION:
; NOTE: 2 variants of devices, with same prod ID but other key codes (however last 4 digits are the same)
usagePage := 65440
vendorID := 2321
productID := 3100

WM_INPUT := 0x00FF
RIDEV_INPUTSINK     := 0x00000100   ;If set, this enables the caller to receive the input even when the caller is not in
DI_HID_VENDORID             := 8    ;Vendor ID for the HID.
DI_HID_PRODUCTID            := 12   ;Product ID for the HID.
II_DEVHANDLE        := 8    ;Handle to the device generating the raw input data.

; Create GUI to receive messages
handlerGUI := Gui()
; handlerGUI.Opt("+LastFound +HwndhandlerGUI")
handlerGUI.Opt("+LastFound")
OnMessage WM_INPUT, InputMsg
; Intercept WM_INPUT messages

; Register Remote Control with RIDEV_INPUTSINK (so that data is received even in the background)
; 65440 and 1 are the usage page and usage, and are unique to the philips device. Use AHKHID_listdevices and AHKHID_interceptdata to get the number of other devices.
; Lower down (in get Input Msg) the vendor and prod id should be entered.
r := AHKHID_Register(usagePage, 1, handlerGUI.Hwnd, RIDEV_INPUTSINK)

; Return ;; TODO: toegevoegd maar misschien breekt dit alles....

funct_caller(funct, param) {			; this will call a the function funct with parameters ==> funct(param)
	funct.call(param)
	; if ( IsFunc(funct) != 0 )
	;	ret_val := %funct%(param)
	; return
}

InputMsg(wParam, lParam, msg, hwnd) {
	;; MsgBox(wParam . " " . lParam . " " . msg . " " . hwnd)
	; 1 905447557 255 40174758
	Local devh, iKey, sLabel
	Critical()

	;Get handle of device
	devh := AHKHID_GetInputInfo(lParam, II_DEVHANDLE)
	; MsgBox(devh . " " . AHKHID_GetDevInfo(devh, DI_HID_VENDORID, True) . " " . AHKHID_GetDevInfo(devh, DI_HID_PRODUCTID, True))

	;Check for error
	If (devh != -1) ;Check that it is the philips device
	And (AHKHID_GetDevInfo(devh, DI_HID_VENDORID, True) = 2321)
	And (AHKHID_GetDevInfo(devh, DI_HID_PRODUCTID, True) = 3100) {

		;Get data
		iKey := AHKHID_GetInputData(lParam, &uData)
		;Check for error
		If (iKey != -1) {
			; udata = adress (buffer)
			; iKey = len
			keyCode := Bin2Hex(&uData, iKey)
			; MsgBox(keyTest)
			keyBuffer(keyCode)
			; PassHotkey(output)			; if you dont want to use the buffer
			; _log(keyCode)
		}
	}
}

keyBuffer(keypressed) {				; Some devices will fire the output twice very fast: this will prevent it from firing twice. The key 5656 (release button) is always passed.
	static lastKeyPressed := keypressed
	if (keypressed == lastKeyPressed and
			SubStr(keypressed, StrLen(keypressed)-3) != "5656") {
			return
	}
	PassHotkey(keypressed)
	lastKeyPressed := keypressed
}

Bin2Hex(&addr, len) {									; magic
    VarSetStrCapacity(&hex, len*(1 ? 2:1)) ; V1toV2: if 'hex' is NOT a UTF-16 string, use 'hex := Buffer(len*(A_IsUnicode ? 2:1))'
    ;; MsgBox(addr.ptr . " len " . addr.size . " len: " . len)
    ; hex := Buffer(len*(2))
; REMOVED:     f := A_FormatInteger
; REMOVED:     SetFormat IntegerFast, H
    Loop len
	hex .= SubStr(0x100 + NumGet(addr.ptr + A_Index-1, 0, "uchar"), -2)
	; hex .= SubStr(0x100 + NumGet(addr + A_Index-1, 0, "uchar"), -2)
; REMOVED:     SetFormat IntegerFast, % f
    return hex
}

/* new keycodes
5688 EOL
5684 i
5620 Ins
5672 Back
5657 Rec
5664 Fw
5660 play
5856 F1
6056 F2
6456 F3
7256 F4
5656 pickup/drop + release button
8856 back button
*/


/*
; Key codes most Philips speech devices
PassHotkey(keypressed) {
	Switch keypressed {
		Case "00800000000000000020":
			MsgBox, EOL
		Case "00800000000000000080":
			MsgBox, -i-
		Case "00800000000000000040":
			MsgBox, insert
		Case "00800000000000000010":
			MsgBox, back
		Case "00800000000000000001":
			MsgBox, record
		Case "00800000000000000008":
			MsgBox, fw
		Case "00800000000000000004":
			MsgBox, Play
		Case "00800000000000000200":
			MsgBox, F1
		Case "00800000000000000400":
			MsgBox, F2
		Case "00800000000000000800":
			MsgBox, F3
		Case "00800000000000001000":
			MsgBox, F4
		Case "00800000000000002000":
			MsgBox, back button
		Case "009E0000000000000000":
			MsgBox, picked up
		Case "009E0000000000000001":
			MsgBox, put down
		Case "00800000000000000000":
			;do nothing, release
		Case keypressed:
			MsgBox, unknown key: %keypressed%
	}
}
*/

/*
; Key codes for some other Philips speech devices.
00800000002404000020	; EOL
00800000002404000080	; -i-
00800000002404000040	; INS
00800000002404000010	; Back
00800000002404000001	; REC
00800000002404000008	; Fwd
00800000002404000004	; Play/pause
00800000002404000200	; F1
009E0000000000000000	; Probably pick up signal (not sure)
00800000002404000400	; F2
00800000002404000800	; F3
00800000002404001000	; F4
00800000002404002000	; BackButton
009E0000000000000001	; Probably put down signal (not sure)
*/
