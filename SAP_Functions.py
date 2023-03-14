def sap_login(environment):
    import win32com.client
    import pythoncom
    import subprocess
    import ctypes
    import sys
    import time
    try:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        proc = subprocess.Popen(path)
        time.sleep(10)

        # Necesario par correr win32com.client en Threading
        pythoncom.CoInitialize()

        sapmin = ctypes.windll.user32.FindWindowW(None, "SAP Logon 780")
        ctypes.windll.user32.ShowWindow(sapmin, 6)

        sapmin = ctypes.windll.user32.FindWindowW(None, "SAP Logon 760")
        ctypes.windll.user32.ShowWindow(sapmin, 6)

        sapmin = ctypes.windll.user32.FindWindowW(None, "SAP Logon 740")
        ctypes.windll.user32.ShowWindow(sapmin, 6)

        sap_gui_auto = win32com.client.GetObject('SAPGUI')
        if not type(sap_gui_auto) == win32com.client.CDispatch:
            return

        application = sap_gui_auto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            sap_gui_auto = None
            return

        if environment == "Q02":
            connection = application.OpenConnection("Q02 - Quality", True)
        if environment == "P02":
            connection = application.OpenConnection("P02 - Production", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            sap_gui_auto = None
            return
        session.ActiveWindow.Iconify()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "SAP_USER"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "SAP_PASSWORD"
        session.findById("wnd[0]").sendVKey(0)
        try:
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

    except:
        print(sys.exc_info())
    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None


def terminate():
    import win32com.client
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not type(sap_gui_auto) == win32com.client.CDispatch:
        return

    application = sap_gui_auto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        sap_gui_auto = None
        return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        sap_gui_auto = None
        return

    if connection.DisabledByServer:
        application = None
        sap_gui_auto = None
        return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        sap_gui_auto = None
        return

    if session.Info.IsLowSpeedConnection:
        connection = None
        application = None
        sap_gui_auto = None
        return
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


# def lt01_query(storage_location, material, quantity, storage_type, storage_bin, destination_storage_type, destination_storage_bin, destination_storage_unit, children):
#     import win32com.client
#     import re
#     import json
#     import time
#     try:
#
#         sap_gui_auto = win32com.client.GetObject("SAPGUI")
#         if not type(sap_gui_auto) == win32com.client.CDispatch:
#             return
#
#         application = sap_gui_auto.GetScriptingEngine
#         if not type(application) == win32com.client.CDispatch:
#             sap_gui_auto = None
#             return
#
#         connection = application.Children(children)
#         if not type(connection) == win32com.client.CDispatch:
#             application = None
#             sap_gui_auto = None
#             return
#
#         if connection.DisabledByServer:
#             application = None
#             sap_gui_auto = None
#             return
#
#         session = connection.Children(0)
#         if not type(session) == win32com.client.CDispatch:
#             connection = None
#             application = None
#             sap_gui_auto = None
#             return
#
#         if session.Info.IsLowSpeedConnection:
#             connection = None
#             application = None
#             sap_gui_auto = None
#             return
#
#         session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT01"
#         session.findById("wnd[0]").sendVKey(0)
#         session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").text = "521"
#         session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "998"
#         session.findById("wnd[0]/usr/ctxtLTAP-MATNR").text = material
#         session.findById("wnd[0]/usr/txtRL03T-ANFME").text = quantity
#         session.findById("wnd[0]/usr/ctxtLTAP-ALTME").text = ""
#         session.findById("wnd[0]/usr/ctxtLTAP-WERKS").text = "5210"
#         session.findById("wnd[0]/usr/ctxtLTAP-LGORT").text = f'00{storage_location}'
#         session.findById("wnd[0]").sendVKey(0)
#         session.findById("wnd[0]/usr/ctxtLTAP-LETYP").text = "001"
#         session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").text = storage_type
#         session.findById("wnd[0]/usr/ctxtLTAP-VLBER").text = "001"
#         session.findById("wnd[0]/usr/txtLTAP-VLPLA").text = storage_bin
#         session.findById("wnd[0]/usr/ctxtLTAP-VLENR").text = destination_storage_unit
#         session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").text = destination_storage_type
#         session.findById("wnd[0]/usr/ctxtLTAP-NLENR").text = destination_storage_unit
#         session.findById("wnd[0]/usr/ctxtLTAP-NLBER").text = "001"
#         session.findById("wnd[0]/usr/txtLTAP-NLPLA").text = destination_storage_bin
#         session.findById("wnd[0]").sendVKey(0)
#         session.findById("wnd[0]").sendVKey(0)
#
#         sap_message = session.findById("wnd[0]/sbar/pane[0]").Text
#         # session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
#         # session.findById("wnd[0]").sendVKey(0)
#
#         time.sleep(.005)
#         try:
#             # Verify if Transfer order in message if not stk.end error state
#             int(re.sub(r"\D", "", sap_message, 0))
#             response = {"result": sap_message, "error": "N/A"}
#             return json.dumps(response)
#         except:
#             response = {"result": "N/A", "error": sap_message}
#             return json.dumps(response)
#
#     except:
#         # print(sys.exc_info()[0])
#         try:
#             session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
#         except:
#             pass
#         time.sleep(.005)
#         # sap_error = session.findById("wnd[0]/sbar/pane[0]").Text
#         session.findById("wnd[0]/tbar[0]/btn[15]").press()
#         # err(sap_error)
#         response = {"result": "N/A", "error": sap_message}
#         return json.dumps(response)
#     finally:
#         session = None
#         connection = None
#         application = None
#         sap_gui_auto = None
#         time.sleep(.005)
#
#
# def ls24_query(sap_num, storage_type, storage_bin, storage_unit, children):
#     import win32com.client
#     import json
#     import pythoncom
#     import re
#     try:
#         pythoncom.CoInitialize()
#         sap_gui_auto = win32com.client.GetObject("SAPGUI")
#
#         application = sap_gui_auto.GetScriptingEngine
#
#         connection = application.Children(children)
#
#         if connection.DisabledByServer:
#             print("Scripting is disabled by server")
#             application = None
#             sap_gui_auto = None
#             return
#
#         session = connection.Children(0)
#
#         if session.Info.IsLowSpeedConnection:
#             print("Connection is low speed")
#             connection = None
#             application = None
#             sap_gui_auto = None
#             return
#
#         session.findById("wnd[0]/tbar[0]/okcd").text = "/nLS24"
#         session.findById("wnd[0]").sendVKey(0)
#         session.findById("wnd[0]/usr/ctxtRL01S-LGNUM").text = "521"
#         session.findById("wnd[0]/usr/ctxtRL01S-MATNR").text = sap_num
#         session.findById("wnd[0]/usr/ctxtRL01S-WERKS").text = "5210"
#         session.findById("wnd[0]/usr/ctxtRL01S-BESTQ").text = "*"
#         session.findById("wnd[0]/usr/ctxtRL01S-SOBKZ").text = "*"
#         session.findById("wnd[0]/usr/ctxtRL01S-LGTYP").text = storage_type
#         session.findById("wnd[0]/usr/ctxtRL01S-LGPLA").text = storage_bin
#         session.findById("wnd[0]/usr/ctxtRL01S-LISTV").text = "/DEL"
#         session.findById("wnd[0]").sendVKey(0)
#
#         try:
#             error = session.findById("wnd[0]/sbar/pane[0]").Text
#             if error != "":
#                 session.findById("wnd[0]/tbar[0]/btn[15]").press()
#                 session.findById("wnd[0]/tbar[0]/btn[15]").press()
#                 response = {"quantity": "N/A", "error": error}
#                 return json.dumps(response)
#             else:
#                 raise Exception('I know Python!')
#         except:
#             try:
#                 quantity = session.findById("wnd[0]/usr/lbl[35,8]").Text
#
#             except:
#                 pass
#             # session.findById("wnd[0]/tbar[0]/btn[15]").press()
#             # session.findById("wnd[0]/tbar[0]/btn[15]").press()
#
#             response = {"quantity": int(float(re.sub(r",", "", quantity).strip())), "error": "N/A"}
#
#             return json.dumps(response)
#
#     except Exception as e:
#         try:
#             session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
#         except:
#             pass
#         error = session.findById("wnd[0]/sbar/pane[0]").Text
#         response = {"quantity": "N/A", "error": error}
#
#         # session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
#         # session.findById("wnd[0]").sendVKey(0)
#         return json.dumps(response)
#
#     finally:
#         session = None
#         connection = None
#         application = None
#         sap_gui_auto = None
#
#
# def lt09_(storage_unit, dst, dsb, children):
#     import pythoncom
#     import win32com.client
#     import json
#     try:
#         pythoncom.CoInitialize()
#         sap_gui_auto = win32com.client.GetObject("SAPGUI")
#
#         application = sap_gui_auto.GetScriptingEngine
#
#         connection = application.Children(children)
#
#         if connection.DisabledByServer:
#             print("Scripting is disabled by server")
#             application = None
#             sap_gui_auto = None
#             return
#
#         session = connection.Children(0)
#
#         if session.Info.IsLowSpeedConnection:
#             print("Connection is low speed")
#             connection = None
#             application = None
#             sap_gui_auto = None
#             return
#
#         session.findById("wnd[0]/tbar[0]/okcd").text = "/nLT09"
#         session.findById("wnd[0]").sendVKey(0)
#         session.findById("wnd[0]/usr/txtLEIN-LENUM").text = storage_unit
#         session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").text = "998"
#         session.findById("wnd[0]").sendVKey(0)
#         # special stock
#         special_stock = session.findById("wnd[0]/usr/subD0171_S:SAPML03T:1711/tblSAPML03TD1711/ctxtLTAP-SOBKZ[10,0]").Text
#         if special_stock == "k" or special_stock == "K":
#             session.findById("wnd[0]/usr/ctxt*LTAP-NLTYP").text = dst
#             session.findById("wnd[0]/usr/ctxt*LTAP-NLBER").text = "001"
#             session.findById("wnd[0]/usr/txt*LTAP-NLPLA").text = dsb
#
#             session.findById("wnd[0]/tbar[0]/btn[11]").press()
#             result = session.findById("wnd[0]/sbar/pane[0]").Text
#         else:
#             response = {"result": "OK", "error": "N/A"}
#             return json.dumps(response)
#         # Getting only the transfer order and not the text
#         session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
#         session.findById("wnd[0]").sendVKey(0)
#
#         try:
#             session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
#         except:
#             pass
#         response = {"result": f'{result}', "error": "N/A"}
#
#         return json.dumps(response)
#
#     except Exception as e:
#         print(e)
#         error = session.findById("wnd[0]/sbar/pane[0]").Text
#         response = {"result": "N/A", "error": error}
#         try:
#             session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
#         except:
#             session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
#             session.findById("wnd[0]").sendVKey(0)
#
#         return json.dumps(response)
#
#     finally:
#         session = None
#         connection = None
#         application = None
#         sap_gui_auto = None


def mb1a_(scrap_material, header, scrap_reason, storage_location, scrap_cost_center, scrap_order, scrap_component, scrap_quantity, children):
    import pythoncom
    import win32com.client
    import json
    try:
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")

        application = sap_gui_auto.GetScriptingEngine

        connection = application.Children(children)

        if connection.DisabledByServer:
            print("Scripting is disabled by server")
            application = None
            sap_gui_auto = None
            return

        session = connection.Children(0)

        if session.Info.IsLowSpeedConnection:
            print("Connection is low speed")
            connection = None
            application = None
            sap_gui_auto = None
            return

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB1A"
        session.findById("wnd[0]").sendVKey(0)
        # session.findById("wnd[0]/usr/chkRM07M-XNAPR").selected = 0
        session.findById("wnd[0]/usr/txtRM07M-MTSNR").text = scrap_material
        session.findById("wnd[0]/usr/txtMKPF-BKTXT").text = header
        session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "551"
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "5210"
        session.findById("wnd[0]/usr/ctxtRM07M-GRUND").text = scrap_reason
        session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = storage_location
        session.findById("wnd[0]/usr/chkRM07M-XNAPR").setFocus()
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").text = scrap_cost_center
        session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-AUFNR").text = scrap_order
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = scrap_component
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = scrap_quantity
        # session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").setFocus()
        # session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").caretPosition = 1
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        result = session.findById("wnd[0]/sbar/pane[0]").Text
        session.findById("wnd[0]").sendVKey(0)


        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
        except:
            pass

        if "Document" not in result:
            response = {"result": "N/A", "error": result}
        else:
            response = {"result": f'{result}', "error": "N/A"}

        return json.dumps(response)

    except Exception as e:
        print(e)
        error = session.findById("wnd[0]/sbar/pane[0]").Text
        response = {"result": "N/A", "error": error}
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
        except:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

        return json.dumps(response)

    finally:
        session = None
        connection = None
        application = None
        sap_gui_auto = None
