import threading
from pywinauto import Application, timings
import time
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys

def mediview_automation():
    try:
        # Start the Mediview application
        app = Application(backend="uia").start(r"C:\Mediview\Mediview.exe")

        # Increase the wait time to 60 seconds and check multiple times
        main_window = None
        for _ in range(60):
            try:
                main_window = app.window(title_re=".*Mediview.*")
                if main_window.exists():
                    break
            except ElementNotFoundError:
                time.sleep(1)

        if main_window is None:
            raise ElementNotFoundError("Mediview window could not be found after 60 seconds.")

        # Bring the main window into focus
        main_window.set_focus()

        # Interact with the "Password" field
        password_field = main_window['Edit2']
        if password_field.exists():
            password_field.type_keys('1', with_spaces=True)  # Replace "1" with your actual password
            print("Mật khẩu đã được nhập.")
            main_window.Button4.click()
            
            # Wait for the interface to update after login
            time.sleep(5)

            # Interact with the "Open Patient" hyperlink
            open_patient_button = main_window.child_window(title="Open Patient", control_type="Hyperlink")
            if open_patient_button.exists():
                open_patient_button.invoke()  # Use invoke to simulate the click
                print("Đã kích hoạt Open Patient.")
                
                # Wait for the interface to update
                time.sleep(5)
                
                # Click on "001" DataItem
                data_item_001 = main_window.child_window(title="001", control_type="DataItem")
                data_item_001.click_input()
                print("Đã click vào mục '001'.")
                
                # Click on "OK" Button
                ok_button = main_window.child_window(title="OK", control_type="Button")
                ok_button.click_input()
                print("Đã click vào nút 'OK'.")
                
                # Wait for any interface update after clicking OK
                time.sleep(5)

                # Attempt to find and click the 'GroupBox5' GroupBox
                select_all_text = main_window.child_window(title="Select All", control_type="Text")
                siblings = select_all_text.parent().children()
                checkbox_nearby = None
                for sibling in siblings:
                    if sibling.element_info.control_type == 'CheckBox':  # Use element_info.control_type to check the control type
                        checkbox_nearby = sibling
                        break
                if checkbox_nearby:
                    checkbox_nearby.click_input()  # Click vào checkbox gần "Select All"
                    print("Đã click vào checkbox gần 'Select All'.")
                    Extern_button = main_window.child_window(title="", control_type="Button")
                    Extern_button.click_input()
                    time.sleep(5)
                    
                    ok_button_2 = main_window.child_window(title=".jpg/.mp4 (HQ)", control_type="RadioButton")
                    ok_button_2.click_input()
                    
                    # Find the "Export path:" Edit field
                    export_path_static = main_window.child_window(title="Exportpath:", control_type="Text")

                    # Tìm tất cả các đối tượng con trong cùng một cấp cha với 'Export path:'
                    all_children = export_path_static.parent().descendants()

                    # Lọc ra đối tượng Edit liên quan đến "Export path:"
                    export_path_edit = None
                    for child in all_children:
                        if child.friendly_class_name() == 'Edit':
                            export_path_edit = child
                            break

                    # Kiểm tra nếu tìm thấy đối tượng Edit và kiểm tra giá trị của nó
                    if export_path_edit:
                        # Lấy giá trị từ trường Edit
                        export_path_value = export_path_edit.get_value()
                        print(f"Giá trị của trường 'Export path:' là: {export_path_value}")

                        # Kiểm tra xem đường dẫn có chứa 'Save_image' hay không
                        if "Save_image" in export_path_value:
                            # Nếu đã có thư mục Save_image, ấn nút OK
                            ok_button_final = main_window.child_window(title="OK", control_type="Button")
                            ok_button_final.click_input()
                            print("Đã tìm thấy thư mục Save_image. Đã nhấn nút OK.")
                        else:
                            # Nếu không có thư mục Save_image, thực hiện các hành động chọn thư mục
                            ok_button_1 = main_window.child_window(title="Browse", control_type="Button")
                            ok_button_1.click_input()

                            # Wait for "Select Folder" window to appear
                            select_folder_window = None
                            for _ in range(30):  # Wait up to 30 seconds
                                try:
                                    select_folder_window = app.window(title="Select Folder")
                                    if select_folder_window.exists():
                                        break
                                except ElementNotFoundError:
                                    time.sleep(1)

                            if select_folder_window is None:
                                raise ElementNotFoundError("Select Folder window could not be found.")

                            # Click on "Save_image"
                            desktop_tree_item = select_folder_window.child_window(title="This PC", control_type="TreeItem")
                            desktop_tree_item.click_input()
                            desktop_tree_item = select_folder_window.child_window(title="Desktop", auto_id="1", control_type="ListItem")
                            desktop_tree_item.double_click_input()
                            select_folder_window.print_control_identifiers()
                            ok_button_final = select_folder_window.child_window(title="New folder", auto_id="{E44616AD-6DF1-4B94-85A4-E465AE8A19DB}", control_type="Button")
                            ok_button_final.click_input()

                            new_folder_edit = select_folder_window.child_window(title="New folder", control_type="Edit")
                            new_folder_edit.type_keys("Save_image", with_spaces=True)
                            send_keys('{ENTER}')
                            ok_button_final = select_folder_window.child_window(title="Select Folder", auto_id="1", control_type="Button")
                            ok_button_final.click_input()
                            time.sleep(5)
                            ok_button_final = main_window.child_window(title="OK", control_type="Button")
                            ok_button_final.click_input()
                            print("Đã hoàn thành việc lưu hình ảnh vào thư mục Save_image.")
                    else:
                        print("Không tìm thấy trường Edit liên quan đến 'Export path:'.")
                else:
                    print("Không tìm thấy checkbox gần 'Select All'.")
                    main_window.print_control_identifiers()
            else:
                print("Không tìm thấy nút 'Open Patient'.")
        else:
            print("Không thể tìm thấy trường mật khẩu.")

    except ElementNotFoundError:
        print("Không thể tìm thấy cửa sổ Mediview sau 60 giây.")
    except Exception as e:
        print(f"Đã xảy ra lỗi: {str(e)}")

def main():
    # Create and start the thread
    mediview_thread = threading.Thread(target=mediview_automation)
    mediview_thread.start()

    # Optionally, wait for the thread to finish
    mediview_thread.join()

if __name__ == "__main__":
    main()
