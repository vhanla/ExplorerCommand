library HotkeyHook;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  System.SysUtils,
  System.Classes,
  Windows,
  Messages;

const
  KeyEvent = WM_USER + 11;
  KeyEventAll = WM_USER + 12;
  KeyEventUpdatePath = WM_USER + 13;
  KeyEventPickPaths = WM_USER + 14;
  LLKHF_ALTDOWN = $20;
  LLKHF_UP = $80;

{ Define a record for recording and passing information process wide }
type
  PKBDLLHOOKSTRUCT = ^TKBDLLHOOKSTRUCT;
  TKBDLLHOOKSTRUCT = record
    vkCode: Cardinal;
    scanCode: Cardinal;
    flags: Cardinal;
    time: Cardinal;
    dwExtrainfo: Cardinal;
  end;

  PHookRec = ^THookRec;
  THookRec = packed record
    HookHandle: HHOOK;
    AppHandle: HWND;
    CtrlWinHandle: HWND;
    KeyCount: DWORD;
    CtrlDown: BOOL;
    ShiftDown: BOOL;
  end;

  TSystemKeyCombination = (skLWin,
  skRWin,
  skCtrlEsc,
  skAltTab,
  skAltEsc,
  skCtrlShiftEsc,
  skAltF4);
  TSystemKeyCombinations = set of TSystemKeyCombination;

{$R *.res}

var
  hObjHandle: THandle; { Variable for the file mappgin object }
  lpHookRec: PHookRec;
  InvalidCombinations: TSystemKeyCombinations;
  AltPressed: BOOL;
  CtrlPressed: BOOL;
  ShiftPressed: BOOL;

procedure SwitchToThisWindow(h1: hWnd; x: bool); stdcall;
  external user32 Name 'SwitchToThisWindow';
{ Pointer to our hook record }
procedure MapFileMemory (dwAllocSize: DWORD);
begin
  { Create a process wide memory mapped variable }
  hObjHandle := CreateFileMapping($FFFFFFFF, nil, PAGE_READWRITE, 0, dwAllocSize, 'AltTabProHook');
  if hObjHandle = 0 then
  begin
    raise Exception.Create('Hook couldn''t create file map object.');
    Exit;
  end;

  { Get a pointer to our process wide memory mapped file }
  lpHookRec := MapViewOfFile(hObjHandle, FILE_MAP_WRITE, 0, 0, dwAllocSize);
  if lpHookRec = nil then
  begin
    CloseHandle(hObjHandle);
    raise Exception.Create('Hook couldn''t map file.');
    Exit;
  end;
end;

procedure UnmapFileMemory;
begin
  { Delete our process wide memory mapped variable }
  if lpHookRec <> nil then
  begin
    UnmapViewOfFile(lpHookRec);
    lpHookRec := nil;
  end;

  if hObjHandle > 0 then
  begin
    CloseHandle(hObjHandle);
    hObjHandle := 0;
  end;
end;

function GetHookRecPointer:Pointer; stdcall;
begin
  { Return a pointer to our process wide memory mapped variable }
  Result := lpHookRec;
end;

function KeyProc(nCode: Integer; wParam: WPARAM; lParam: LPARAM): LRESULT; stdcall;
begin

  /// If code is less than zero, the hook procedure must pass the message
  ///  to the CallNextHookEx function without further processing and should return
  ///  the value returned by CallNextHookEx.

  if nCode < 0 then
  begin
    Result := CallNextHookEx(lpHookRec^.HookHandle, nCode, wParam, lParam);
    Exit;
  end;

  /// HC_ACTION = 0 : The wParam and lParam parameters contain information
  ///  about a keystroke message.

  if nCode = HC_ACTION then
  begin

  end

  /// HC_NOREMOVE = 3 : The wParam and lParam parameters contain information
  ///  about a keystroke message, and the keystroke message has not been removed
  ///  from the message queue. (An application called the PeekMessage function,
  ///  specifying the PM_NOREMOVE flag).
  else if nCode = HC_NOREMOVE then
  begin

  end;





end;

function KeyboardProc(nCode: Integer; wParam: WPARAM; lParam: LPARAM): LRESULT; stdcall;
var
  KeyUp: BOOL;
//  AltPressed: BOOL;
//  CtrlPressed: BOOL;
//  ShiftPressed: BOOL;
  KeyName: string;
  Res: Integer;
  ParentHandle: HWND;
  hs: PKBDLLHOOKSTRUCT;
  command: string;
  AppClassName: array[0..255] of char;
  currWnd: HWND;
begin
  Result := 0;

  case nCode of
    HC_ACTION: // HC_ACTION is the only allowed for WH_KEYBOARD_LL
    begin

      hs := PKBDLLHOOKSTRUCT(lParam);

      if (wParam = WM_KEYDOWN) or (wParam = WM_SYSKEYDOWN) then
      begin
        if (hs^.vkCode = VK_MENU) or (hs^.vkCode = VK_LMENU) or (hs^.vkCode = VK_RMENU) then
          AltPressed := True;

//        OutputDebugString(PChar(IntToStr(hs^.vkcode)));
        if (hs^.vkCode = VK_SHIFT) or (hs^.vkCode = VK_LSHIFT) or (hs^.vkCode = VK_RSHIFT) then
        begin
//          lpHookRec^.CtrlDown := True;
          ShiftPressed := True;
//          OutputDebugString('ShiftPressed');

        end;

        if (hs^.vkCode = VK_CONTROL) or (hs^.vkCode = VK_LCONTROL) or (hs^.vkCode = VK_RCONTROL) then
        begin
//          lpHookRec^.ShiftDown := True;
          CtrlPressed := True;
//          OutputDebugString('CtrlPressed');
        end;
      end;

      if (wParam = WM_KEYUP) or (wParam = WM_SYSKEYUP) then
      begin

        /// NOTE: When this callback function is called in response to a change
        ///  in the state of a key, the callback function is called before the
        ///  asynchronous state of the key is updated. Consequently, the
        ///  asynchronous state of the key cannot be determined by calling
        ///  GetAsyncKeyState from within this callback
        ///  HOWEVER, this works :P
//        CtrlPressed := GetAsyncKeyState(VK_CONTROL) and $8000 <> 0;
//        ShiftPressed := GetAsyncKeyState(VK_SHIFT) and $8000 <> 0;
//        AltPressed := GetAsyncKeyState(VK_MENU) and $8000 <> 0;
        if (hs^.vkCode = VK_MENU) or (hs^.vkCode = VK_LMENU) or (hs^.vkCode = VK_RMENU) then
          AltPressed := False;

        if (hs^.vkCode = VK_SHIFT) or (hs^.vkCode = VK_LSHIFT) or (hs^.vkCode = VK_RSHIFT) then
        begin
//          lpHookRec^.CtrlDown := False;
          ShiftPressed := False;
//          OutputDebugString('CtrlUnpressed');
        end;

        if (hs^.vkCode = VK_CONTROL) or (hs^.vkCode = VK_LCONTROL) or (hs^.vkCode = VK_RCONTROL) then
        begin
//          lpHookRec^.ShiftDown := False;
          CtrlPressed := False;
//          OutputDebugString('ShiftUnpressed');
        end;

//########### Actual Hotkey overrides assignments ###########
        if ((hs^.vkCode = Ord('P')) and ShiftPressed and CtrlPressed) then
//        if ((hs^.vkCode = Ord('P')) and CtrlPressed and (hs^.flags and LLKHF_ALTDOWN = 1)) then
//        if ((hs^.vkCode = Ord('P')) and lpHookRec^.CtrlDown and lpHookRec^.ShiftDown) then
//        if hs^.vkCode = VK_SPACE then
        begin

          currWnd := GetForegroundWindow;
          if currWnd > 0 then
          begin
            GetClassName(currWnd, AppClassName, 255);
            // Let's also show when IFileDialog variants are shown (Open Save Dialog)
            var isFileDialog := FindWindowEx(currWnd, 0, 'DUIViewWndClassName', nil);
            if isFileDialog > 0 then
              isFileDialog := FindWindowEx(isFileDialog, 0, 'DirectUIHWND', nil);

            if (AppClassName <> 'CabinetWClass') then
            begin
              if isFileDialog <= 0 then
              begin
              Result := CallNextHookEx(lpHookRec^.HookHandle, nCode, wParam, lParam);
              Exit;
              end;
            end;
          end;

          ParentHandle := FindWindow('ExplorerCommandWnd', nil);
          if ParentHandle > 0 then
          begin

            //if ShiftPressed then command := 'prev' else command := 'next';
            command := IntToStr(GetForegroundWindow);

            //if (hs^.flags and LLKHF_UP) <> 0 then

            /// The hook procedure should process a message in less time than the data entry specified in the LowLevelHooksTimeout value in the following registry key:
            ///  HKEY_CURRENT_USER\Control Panel\Desktop
            ///  The value is in milliseconds. If the hook procedure times out, the system passes the message to the
            ///  next hook. However, on Windows 7 and later, the hook is silently removed without being called.
            ///  There is no way for the application to know whether the hook is removed.
            //SendMessageTimeout(ParentHandle, KeyEvent, wParam, Windows.LPARAM(PChar(command)), SMTO_NORMAL, 500, nil);
            PostMessage(ParentHandle, KeyEvent, wParam, Windows.LPARAM(PChar(command)));

//            if GetForegroundWindow <> ParentHandle then
//            begin
//              //ShowWindow(ParentHandle, SW_SHOWNORMAL);
//              SetForegroundWindow(ParentHandle);
//              SwitchToThisWindow(ParentHandle, True);
//            end;

//            Exit(1);

          end;

        end
//        VK_OEM_1: Used for the 'ñÑ' key.
//        VK_OEM_2: Used for the 'çÇ' key.
//        VK_OEM_3: Used for the 'ºª' key.
//        VK_OEM_4: Used for the '¡!' key.
//        VK_OEM_5: Used for the '¿?' key.
//        VK_OEM_6: Used for the '´¨' key.
//        VK_OEM_7: Used for the '+*' key.
//        VK_OEM_8: Not commonly used on a Spanish keyboard.
//        VK_OEM_102: Used for the '<>' key, located between the left Shift and Z keys.
//        else if ((hs^.vkCode = Ord('I')) and ShiftPressed and CtrlPressed) then
        else if ((hs^.vkCode = VK_OEM_2 )
        and ShiftPressed and CtrlPressed)
        then
        begin
          ParentHandle := FindWindow('ExplorerCommandWnd', nil);
          if ParentHandle > 0 then
          begin
            command := IntToStr(GetForegroundWindow);
            PostMessage(ParentHandle, KeyEventAll, wParam, Windows.LPARAM(PChar(command)));
          end;
        end

        // Ctrl+Alt+UpArrow if Open Dialog or Save Dialog window detected to assign last explorer's path
        else if ((hs^.vkCode = VK_UP)
        and AltPressed and CtrlPressed)
        then
        begin
          ParentHandle := FindWindow('ExplorerCommandWnd', nil);
          if ParentHandle > 0 then
          begin
            command := IntToStr(GetForegroundWindow);
            PostMessage(ParentHandle, KeyEventUpdatePath, wParam, Windows.LPARAM(PChar(command)));
          end;
        end
        // Ctrl+Alt+DownArrow it will show a list of custom paths to pick either to assign to explorer's current path or open/save dialog even to open in custom app like terminal
        else if ((hs^.vkCode = VK_DOWN)
        and AltPressed and CtrlPressed)
        then
        begin
          ParentHandle := FindWindow('ExplorerCommandWnd', nil);
          if ParentHandle > 0 then
          begin
            command := IntToStr(GetForegroundWindow);
            PostMessage(ParentHandle, KeyEventPickPaths, wParam, Windows.LPARAM(PChar(command)));
          end;
        end;


//        if (hs^.vkCode = VK_TAB) and ((hs^.flags and LLKHF_UP) <> 0) then
//        begin
          //ShowWindow(ParentHandle, SW_HIDE);
//        end;
      end;
//      Result := CallNextHookEx(lpHookRec^.HookHandle, nCode, wParam, lParam);
    end;


    // There is no HC_NOREMOVE in WH_KEYBOARD_ll
//    HC_NOREMOVE:
//    begin
      { This is a keystroke message, but the keystroke message has not been
      removed from the message queue, since an application has called
      PeekMessage() specifying PM_NOREMOVE }
//      Result := 0;
//      Exit;
//    end;
  end;

// as we are not blocking for other hooks/system to process this keyboard hook let's just pass all
//  if nCode < 0 then
  Result := CallNextHookEx(lpHookRec^.HookHandle, nCode, wParam, lParam);
end;

function StartHook:BOOL stdcall;
begin
  Result := False;
  { If we have a process wide memory variable and the hook has not already bee set }
  if ((lpHookRec <> nil) and (lpHookRec^.HookHandle = 0)) then
  begin
    { Set the hook and remember our hook handle }
    lpHookRec^.HookHandle := SetWindowsHookEx(WH_KEYBOARD_LL, @KeyboardProc, HInstance, 0);
    Result := True;
  end;
end;

procedure StopHook; stdcall;
begin
  { If we have a process wide memory variable and the hook has already been ser }
  if ((lpHookRec <> nil) and (lpHookRec^.HookHandle <> 0)) then
  begin
    { Remove our hook and clear our hook handle }
    if (UnhookWindowsHookEx(lpHookRec^.HookHandle) <> False) then
    begin
      lpHookRec^.HookHandle := 0;
    end;
  end;
end;

procedure DllEntryPoint(dwReason: DWORD);
begin
  case dwReason of
    DLL_PROCESS_ATTACH:
    begin
      { If we are getting mapped into a process, then get a pointer
        to our process wide memory mapped variable }
      hObjHandle := 0;
      lpHookRec := nil;
      MapFileMemory(SizeOf(lpHookRec^));
    end;
    DLL_PROCESS_DETACH:
    begin
      { If we are getting unmapped from a proces then, remove the
        pointer to our process wide memory mapped variable }
      UnmapFileMemory;
    end;
  end;
end;

Exports
  KeyboardProc Name 'KEYBOARDPROC',
  GetHookRecPointer name 'GETHOOKRECPOINTER',
  StartHook name 'STARTHOOK',
  StopHook name 'STOPHOOK';

begin
  DllProc := @DllEntryPoint;
  DllEntryPoint(DLL_PROCESS_ATTACH);
end.

