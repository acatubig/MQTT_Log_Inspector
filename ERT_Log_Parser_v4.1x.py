import tkinter as tk
from tkinter import ttk
from tkinter import font
import paho.mqtt.client as mqtt
from datetime import datetime
import json
import base64
import babel.numbers
import binascii
import re
import os
import pandas as pd
from tkcalendar import Calendar
import threading
import time
from tkinter import messagebox

from tkinter.scrolledtext import ScrolledText

class MainWindow:
    
    def __init__(self, master):
        # set background color and size of root window
        master.config(bg="#555")
        master.geometry("1300x700")
        master.minsize(1300,700)
        master.maxsize(1300,700)

        # create a Calendar widget
        self.cal = Calendar(master, selectmode='day', date_pattern='yyyy-mm-dd')
        self.cal.place(relx=1.0, x=-10, y=34, anchor='ne')

        # create a clock widget
        self.clock = tk.Label(master, text=datetime.now().strftime("%I:%M:%S %p"), bg="black", fg="red", font=("LCD", 9))
        self.clock.place(relx=1.0, x=-10, y=5, anchor='ne')

        # Create a font object with desired font family, size, and style and Set the background and foreground colors for the listbox
        custom_font = font.Font(family="Helvetica", size=10, weight="bold")
        background_color = "lightblue"
        foreground_color = "navy blue"

        # add treeview to the frame
        # create a treeview widget with columns and headings
        # set up the treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview.Heading", background="#D3D3D3", fieldbackground="#D3D3D3", font=("Arial", 10, "bold"))
        self.tree = ttk.Treeview(master, style="Custom.Treeview")
        self.tree["columns"] = ("Timestamp", "ERT ID", "Consumption", "Device Name", "DevEUI", "ERT Data", "RSSI", "SNR", "Time Diff")
        self.tree.heading("#0", text="Message")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.heading("Timestamp", text="Timestamp", anchor=tk.W, command=lambda: self.sort_column("Timestamp", False))
        self.tree.column("Timestamp", anchor=tk.W, width=80)
        self.tree.heading("ERT ID", text="ERT ID", anchor=tk.W, command=lambda: self.sort_column("ERT ID", False))
        self.tree.column("ERT ID", anchor=tk.W, width=50)
        self.tree.heading("Consumption", text="Consumption", anchor=tk.W, command=lambda: self.sort_column("Consumption", False))
        self.tree.column("Consumption", anchor=tk.W, width=50)
        self.tree.heading("Device Name", text="Device Name", anchor=tk.W, command=lambda: self.sort_column("Device Name", False))
        self.tree.column("Device Name", anchor=tk.W, width=50)
        self.tree.heading("DevEUI", text="DevEUI", anchor=tk.W, command=lambda: self.sort_column("DevEUI", False))
        self.tree.column("DevEUI", anchor=tk.W, width=80)
        self.tree.heading("ERT Data", text="ERT Data", anchor=tk.W, command=lambda: self.sort_column("ERT Data", False))
        self.tree.column("ERT Data", anchor=tk.W, width=300)
        self.tree.heading("RSSI", text="RSSI", anchor=tk.W, command=lambda: self.sort_column("RSSI", False))
        self.tree.column("RSSI", anchor=tk.W, width=10)
        self.tree.heading("SNR", text="SNR", anchor=tk.W, command=lambda: self.sort_column("SNR", False))
        self.tree.column("SNR", anchor=tk.W, width=10)
        self.tree.heading("Time Diff", text="Time Diff", anchor=tk.W, command=lambda: self.sort_column("Time Diff", False))
        self.tree.column("Time Diff", anchor=tk.W, width=10)
        self.tree.grid(row=6, column=0, columnspan=8, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self.find_repeating_ert_id, self.find_repeating_device)

        # set up GUI components
        self.host_label = tk.Label(master, text="MQTT Broker Host :", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.host_options = ["sample1 - 1.0.0.1","sample2 - 1.0.0.2"]
        self.host_var = tk.StringVar(value=self.host_options[0])
        self.host_menu = tk.OptionMenu(master, self.host_var, *self.host_options)
        self.host_menu.config(font=("Open Sans", 10), width=14)
        self.port_label = tk.Label(master, text="MQTT Broker Port :", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.port_edit = tk.Entry(master, font=("Open Sans", 12),width=15)
        self.port_edit.insert(0, "1883")
        self.connect_button = tk.Button(master, text="CONNECT", command=self.connect_to_broker, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12,"bold"), width=29)
        self.topic_label = tk.Label(master, text="Subscribe to Topic :", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.topic_edit = tk.Entry(master, font=("Open Sans", 12),width=15)
        self.topic_edit.insert(0, "application/1/device/#")
        self.subscribe_button = tk.Button(master, text="SUBSCRIBE", command=self.subscribe_to_topic, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12,"bold"), width=29)
        self.message_label = tk.Label(master, text="Messages:", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.file_location_label = tk.Label(master, text="Save Excel Location:", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.file_location_entry = tk.Entry(master, font=("Open Sans", 12), width=50)
        self.file_location_entry.insert(0, "C:/Users/ADMIN/Desktop/sample_log/MQTT_Data.xlsx")
        self.text_location_label = tk.Label(master, text="Save Text Location:", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.text_location_entry = tk.Entry(master, font=("Open Sans", 12), width=50)
        self.text_location_entry.insert(0, "C:/Users/ADMIN/Desktop/sample_log/ERT_Data.xlsx")
        self.exportexcel_button = tk.Button(master, text="Export ERT Messages to Excel", command=self.export_to_excel, bd=3, bg="#555", fg="#fff", font=("Open Sans", 12),width=65)
        self.exportdevices_button = tk.Button(master, text="Export ERT IDS and Meters to Excel", command=self.export_devices_excel, bd=3, bg="#555", fg="#fff", font=("Open Sans", 12),width=65)
        self.ertidcounter_label = tk.Label(master, text="ERT ID Count:", bg="#555", fg="#fff", font=("Open Sans", 9))
        self.ertidcounter_label.place(relx=1.0, x=-180, y=6, anchor='ne')
        self.ertidcounter_listbox = tk.Listbox(master, height=1, width=10, font=custom_font, bg=background_color, fg=foreground_color)
        self.ertidcounter_listbox.place(relx=1.0, x=-105, y=6, anchor='ne')

        # new

        self.jsonmessage_label = tk.Label(master, text="JSON Message:", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.jsonmessage_entry = ScrolledText(master, wrap=tk.WORD, width=10, height=10, font=("Open Sans", 12))
        self.topic_publish_label = tk.Label(master, text="MQTT Topic to Publish:", bg="#555", fg="#fff", font=("Open Sans", 12))
        self.topic_publish_entry = tk.Entry(master, font=("Open Sans", 12), width=53) #{"data":"jao=", "fport":2}
        self.topic_publish_entry.insert(0, "application/1/device/devEUI_here/tx")
        self.publish_button = tk.Button(master, text="PUBLISH", command=self.publish_message, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12,"bold"), width=47)
        
        # end new

        # create vertical scrollbar for incoming messages
        messages_scrollbar = ttk.Scrollbar(master, orient=tk.VERTICAL, command=self.tree.yview)
        messages_scrollbar.place(relx=1.0, x=-2, y=385, anchor='e', height=255)
        self.tree.config(yscrollcommand=messages_scrollbar.set)

        # create listbox for repeating ERT IDs
        self.ert_listbox = tk.Listbox(master, height=13, width=57)
        self.ert_listbox.place(relx=1.0, x=-270, y=10, anchor='ne')
        self.ert_listbox.insert(tk.END, "Repeating ERT IDs and Corresponding Device Name:")

        # create vertical scrollbar for ert_listbox
        vert_scrollbar = tk.Scrollbar(master, orient=tk.VERTICAL, command=self.ert_listbox.yview)
        vert_scrollbar.place(relx=1.0, x=-270, y=10, anchor='ne', height=210)
        self.ert_listbox.config(yscrollcommand=vert_scrollbar.set)

        # create horizontal scrollbar for ert_listbox
        horz_scrollbar = tk.Scrollbar(master, orient=tk.HORIZONTAL, command=self.ert_listbox.xview)
        horz_scrollbar.place(relx=1.0, x=-285, y=205, anchor='ne', width= 328)
        self.ert_listbox.config(xscrollcommand=horz_scrollbar.set)

        # create listbox for repeating Device Names
        self.devname_listbox = tk.Listbox(master, height=13, width=57)
        self.devname_listbox.place(relx=1.0, x=-630, y=10, anchor='ne')
        self.devname_listbox.insert(tk.END, "Repeating Device and Corresponding ERT ID:")
    
        # create vertical scrollbar for repeating Device Names
        vert_scrollbar = tk.Scrollbar(master, orient=tk.VERTICAL, command=self.devname_listbox.yview)
        vert_scrollbar.place(relx=1.0, x=-630, y=10, anchor='ne', height=210)
        self.devname_listbox.config(yscrollcommand=vert_scrollbar.set)

        # create horizontal scrollbar for repeating Device Names
        horz_scrollbar = tk.Scrollbar(master, orient=tk.HORIZONTAL, command=self.devname_listbox.xview)
        horz_scrollbar.place(relx=1.0, x=-645, y=205, anchor='ne', width= 328)
        self.devname_listbox.config(xscrollcommand=horz_scrollbar.set)

        # set up layout
        self.host_label.grid(row=0, column=0)
        self.host_menu.grid(row=0, column=1, sticky="w", pady=10)
        self.port_label.grid(row=1, column=0)
        self.port_edit.grid(row=1, column=1, sticky="w", pady=10)
        self.connect_button.grid(row=2, column=0, columnspan=5, padx=10, sticky="w")
        self.topic_label.grid(row=3, column=0)
        self.topic_edit.grid(row=3, column=1, sticky="w", pady=10)
        self.subscribe_button.grid(row=4, column=0, columnspan=5, padx=10, sticky="w")
        self.message_label.grid(row=5, column=0,sticky="w", padx=10)

        self.text_location_label.grid(row=7, column=0, sticky="w", padx=10, pady=5)
        self.text_location_entry.grid(row=7, column=1, sticky="w", padx=10, pady=5)
        self.exportexcel_button.grid(row=8, column=0, columnspan=8, sticky="w", padx=10, pady=5)        
        self.file_location_label.grid(row=9, column=0, sticky="w", padx=10, pady=10)
        self.file_location_entry.grid(row=9, column=1, sticky="w", padx=10, pady=10)
        self.exportdevices_button.grid(row=10, column=0, columnspan=8, sticky="w", padx=10, pady=5)
        
        self.topic_publish_label.grid(row=7, column=2, sticky="e", padx=10, pady=5)
        self.topic_publish_entry.grid(row=7, column=3, sticky="w", padx=10, pady=5)
        self.jsonmessage_label.grid(row=8, column=2, sticky="w", padx=10, pady=5)
        self.publish_button.grid(row=8, column=3, columnspan=5, sticky="w", padx=10, pady=5)
        self.jsonmessage_entry.grid(row=9, column=2, columnspan=2, rowspan=2, sticky="nsew", padx=10, pady=10)
        self.jsonmessage_entry.pack_propagate(False)
        
        master.rowconfigure(6, weight=1)
        master.columnconfigure(1, weight=1)

        # create MQTT client instance
        self.client = mqtt.Client()

        # register callback functions for connection and message events
        self.client.on_connect = self.on_connect
        self.client.on_disconnect = self.on_disconnect
        self.client.on_message = self.on_message

        # define incoming_messages attribute
        self.incoming_messages = []
        self.incoming_ertid = []
        self.unique_ert_ids = set()
        self.payload_counter = 0
        self.prev_ert_timestamp = None

        master.after(0, self.update_clock)

    def publish_message(self):
        # Get the JSON message and MQTT topic from the input fields
        json_message = self.jsonmessage_entry.get("1.0", "end-1c")
        mqtt_topic = self.topic_publish_entry.get()

        # Publish the JSON message to the MQTT topic
        try:
            self.client.publish(mqtt_topic, json_message)
            messagebox.showinfo("Message Published", "JSON message published successfully.")
        except Exception as e:
            messagebox.showerror("Publish Error", f"Error publishing message: {e}")
    
    def update_clock(self):
        current_time = datetime.now().strftime("%I:%M:%S %p")
        self.clock.config(text=current_time)
        self.clock.after(1000, self.update_clock)
        
    def on_connect(self, client, userdata, flags, rc):
        # display connection status as pop-up message
        # Check if the connection is a reconnection
        is_reconnection = False
        if userdata:
            is_reconnection = True

        if rc == 0:
            if is_reconnection:
                self.subscribe_to_topic()
                # toggle button to disconnect
                self.connect_button.config(text="DISCONNECT", command=self.disconnect_from_broker, bd=3, bg="limegreen", fg="black", font=("Open Sans", 12, "bold"), width=29)
                #messagebox.showinfo("Reconnected", "Reconnected to MQTT broker. Subscribing to topic...")
                print("Reconnected to MQTT broker.")
            else:
                messagebox.showinfo("Connection Status", "Connected to MQTT broker.")
                # toggle button to disconnect
                self.connect_button.config(text="DISCONNECT", command=self.disconnect_from_broker, bd=3, bg="limegreen", fg="black", font=("Open Sans", 12, "bold"), width=29)
                print("Connected to MQTT broker.")
        else:
            messagebox.showwarning("Connection Status", f"Connection failed with error code {rc}.")
        

    def reconnect(self):
        # reconnect to MQTT broker
        self.client.reconnect()
        print("Reconnected!")

        # Set the userdata to indicate it's a reconnection
        self.client.user_data_set(True)


    def connect_to_broker(self):
        # set up on_disconnect callback
        self.client.on_disconnect = self.on_disconnect

        # Clear the userdata to indicate it's a fresh connection
        self.client.user_data_set(None)

        # get broker host and port from GUI
        host = self.host_var.get().split(" - ")[1]
        port = int(self.port_edit.get())

        # connect to MQTT broker
        self.client.connect(host, port)
        self.client.loop_start()

 
    def on_disconnect(self, client, userdata, rc):
        if rc == 0:
            # if rc is 0, it means the disconnection was intentional, so no need to show a pop-up message
            self.connect_button.config(text="CONNECT", command=self.connect_to_broker, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12, "bold"), width=29)
            self.subscribe_button.config(text="SUBSCRIBE", command=self.subscribe_to_topic, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12, "bold"), width=29)
            print("Disconnected from the broker.")
        else:
            # if rc is not 0, it means there was an error and the disconnection was unexpected, so show a pop-up message
            self.connect_button.config(text="CONNECT", command=self.connect_to_broker, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12, "bold"), width=29)
            self.subscribe_button.config(text="SUBSCRIBE", command=self.subscribe_to_topic, bd=3, bg="#D3D3D3", fg="black", font=("Open Sans", 12, "bold"), width=29)
            print(f"Disconnected from the broker with error code {rc}.")
            while True:
                try:
                    self.reconnect()
                    print("trying to reconnect...")
                    break
                except:
                    print("Reconnecting...")
    
    def disconnect_from_broker(self):
        # toggle button to connect
        try:
            # disconnect from MQTT broker
            self.client.loop_stop()
            time.sleep(1)
            self.client.disconnect()
        except Exception as e:
            # Handle the exception here
            print(f"Error occurred during disconnection: {e}")
        
    def subscribe_to_topic(self):
        # get topic to subscribe from GUI
        topic = self.topic_edit.get()

        # subscribe to MQTT topic
        self.client.subscribe(topic)
        #self.message_listbox.insert(tk.END, "Subscribed to " + topic)
        self.subscribe_button.config(text="SUBSCRIBED", command=self.subscribe_to_topic, bd=3, bg="limegreen", fg="black", font=("Open Sans", 12, "bold"), width=29)
        print("Subscribed to topic!")
                 
    def on_message(self, client, userdata, message):
        # extract specific fields from incoming message and display in GUI
        payload = message.payload.decode()
        try:
            data = json.loads(payload)
            device_name = data.get("deviceName", "")
            device_eui = data.get("devEUI", "")
            data_value = data.get("data", "")
            rx_info = data.get("rxInfo", [])
            if rx_info:
                rssi = rx_info[0].get("rssi")
                lora_snr = rx_info[0].get("loRaSNR")
                #print(f"RSSI: {rssi}, LoRaSNR: {lora_snr}")
            hex_value = ""
            if data_value:
                decoded_data = base64.b64decode(data_value)
                hex_value = decoded_data.hex()

            if hex_value.startswith("8e"):
                # Display device_name, dev_eui, and converted data_value in a pop-up message
                message_text = f"Device Name: {device_name}\nDevEUI: {device_eui}\nEchoed Data Value (Hex): {hex_value}"
                messagebox.showinfo("Echoed Data Information", message_text)

            if hex_value.startswith("9e"):
                ert_id = int(hex_value[30:38], 16)
                if ert_id not in self.unique_ert_ids:
                    self.payload_counter += 1
                    self.unique_ert_ids.add(ert_id)
                consumption_data = int(hex_value[8:16], 16)
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                current_timestamp = datetime.now()
                if self.prev_ert_timestamp is not None:
                    time_diff = (current_timestamp - self.prev_ert_timestamp).total_seconds()
                else:
                    time_diff = None
                self.prev_ert_timestamp = current_timestamp
                self.incoming_messages.append((current_time, ert_id, consumption_data, device_name, device_eui, hex_value, rssi, lora_snr, time_diff))
                # add data to the treeview
                self.tree.tag_configure("CustomFont",foreground='navy blue', font=("Calibri", 10, "bold"))
                self.tree.insert(parent="", index="end", text="", values=(current_time, ert_id, consumption_data, device_name, device_eui, hex_value, rssi, lora_snr, time_diff), tags=("CustomFont",))
                self.tree.yview_moveto(1.0)
                print(time_diff)

                # Create separate threads for each function
                thread_ert_id = threading.Thread(target=self.find_repeating_ert_id)
                thread_device = threading.Thread(target=self.find_repeating_device)
                thread_count = threading.Thread(target=self.find_ert_id_count)

                # Start the threads
                thread_ert_id.start()
                thread_device.start()
                thread_count.start()

                '''self.find_repeating_ert_id()
                self.find_repeating_device()
                self.find_ert_id_count()'''

        except json.JSONDecodeError:
            message_text = payload
            #self.message_listbox.insert(tk.END, message_text)
            #self.message_listbox.see(tk.END) # scroll to the bottom of the listbox
    
    def find_ert_id_count(self):
        self.ertidcounter_listbox.delete(0, tk.END)
        self.ertidcounter_listbox.insert(0, self.payload_counter)

    def find_repeating_ert_id(self):
        # create a dictionary to store the count of each ERT ID and associated device names
        ert_counts = {}
        for msg in self.incoming_messages:
            ert_id = msg[1]
            devname = msg[3]
            if devname:
                device_name = devname
            else:
                device_name = "Unknown"
            if ert_id in ert_counts:
                ert_counts[ert_id]["count"] += 1
                ert_counts[ert_id]["devices"].add(device_name)
            else:
                ert_counts[ert_id] = {"count": 1, "devices": set([device_name])}

        # create a list of ERT IDs that occur more than once
        repeating_ert_ids = [key for key, value in ert_counts.items() if value["count"] > 1]

        # clear the listbox and insert new data
        self.ert_listbox.delete(1, tk.END)

        for ert_id in repeating_ert_ids:
            count = ert_counts[ert_id]["count"]
            devices = ", ".join(sorted(ert_counts[ert_id]["devices"]))
            self.ert_listbox.insert(1, f"ERT ID: {ert_id} ({count} times) - Devices: {devices}")
        
        self.ert_listbox.yview(tk.END)

    def find_repeating_device(self):
        # create a dictionary to store the count of each Device and associated ERT ID
        dev_counts = {}
        for msg in self.incoming_messages:
            devname = msg[3]
            ert_id = msg[1]
            if ert_id:
                ert_name = ert_id
            else:
                ert_name = "Unknown"
            if devname in dev_counts:
                dev_counts[devname]["count"] += 1
                dev_counts[devname]["ert"].add(ert_name)
            else:
                dev_counts[devname] = {"count": 1, "ert": set([ert_name])}

        # create a list of Device Name that occur more than once
        repeating_dev_name = [key for key, value in dev_counts.items() if value["count"] > 1]

        # clear the listbox and insert new data
        self.devname_listbox.delete(1, tk.END)

        for devname in  repeating_dev_name:
            count = dev_counts[devname]["count"]
            erts = sorted([int(ert) for ert in dev_counts[devname]["ert"] if ert != 'Unknown'])
            
            self.devname_listbox.insert(1, f"Device Name: {devname} ({count} times) - ERT ID: {erts}")
        
        self.devname_listbox.yview(tk.END)

    def export_to_excel(self):
        # create a pandas dataframe from incoming_messages list
        df = pd.DataFrame(self.incoming_messages, columns=["Timestamp", "ERT ID", "Consumption", "Device Name", "DevEUI", "ERT Data", "RRSI", "SNR", "Time Diff"])
        
        # create an Excel writer object
        file_location = self.file_location_entry.get()
        writer = pd.ExcelWriter(file_location)
        
        # write dataframe to Excel file
        df.to_excel(writer, index=False)
        
        # save the Excel file and close the writer object
        writer._save()

    def export_devices_excel(self):
        # create an empty list to hold devices for each ERT ID
        devices_list = []

        # iterate through each item in the listbox
        for i in range(self.ert_listbox.size()):
            # get the text of the listbox item
            item = self.ert_listbox.get(i)

            # extract ERT ID, count, and devices from the text
            ert_id = item.split(" ")[2]
            count = item.split(" ")[4]
            devices = item.split(": ")[-1].split(", ")

            # strip whitespace from device numbers and append to devices list
            devices = [device.strip() for device in devices]
            devices_list.append(devices)

        # determine maximum number of devices
        max_devices = max(len(devices) for devices in devices_list)

        # create an empty dataframe with columns
        columns = ['ERT ID', 'Count'] + [f'Device {i+1}' for i in range(max_devices)]
        df = pd.DataFrame(columns=columns)

        # iterate through each item in the listbox
        for i in range(1,self.ert_listbox.size()):
            # get the text of the listbox item
            item = self.ert_listbox.get(i)

            # extract ERT ID, count, and devices from the text
            match = re.search(r'\((\d+) times\)', item)
            ert_id = item.split(" ")[2]
            count = match.group(1)
            devices = devices_list[i]

            # create a new row in the dataframe with the extracted data
            row = [ert_id, count] + ["" for _ in range(max_devices)]
            for j, device in enumerate(devices):
                row[j+2] = device.strip()
            df.loc[i] = row

        # save the dataframe to Excel
        file_location = self.text_location_entry.get()
        #self.message_listbox.insert(tk.END, f"Data Successfully Saved to {file_location}")
        df.to_excel(file_location, index=False)
        self.export_erts_excel()

    def export_erts_excel(self):
        # create an empty list to hold ERT ID for each devices
        ert_list = []

        # iterate through each item in the listbox
        for i in range(1,self.devname_listbox.size()):
            # get the text of the listbox item
            item = self.devname_listbox.get(i)

            # extract ERT ID and count from the text
            erts_str = item.split("[")[1].split("]")[0]
            ertid = erts_str.split(", ")
            count = item.split(" ")[4]

            # strip whitespace from device numbers and append to devices list
            ertid = [ert.strip() for ert in ertid]
            ert_list.append(ertid)

        # determine maximum number of ERT IDs
        max_erts = max(len(erts) for erts in ert_list)

        # create an empty dataframe with columns
        columns = ['Device Name', 'Count'] + [f'ERT ID {i+1}' for i in range(max_erts)]
        df = pd.DataFrame(columns=columns)

        # iterate through each item in the listbox
        for i in range(1,self.devname_listbox.size()):
            # get the text of the listbox item
            item = self.devname_listbox.get(i)

            # extract ERT ID, count, and devices from the text 
            device_name = item.split(" ")[2]
            count = item.split("(")[1].split(" ")[0]
            erts = ert_list[i-1]

            # create a new row in the dataframe with the extracted data
            row = [device_name, count] + ["" for _ in range(max_erts)]
            for j, ert in enumerate(erts):
                row[j+2] = ert.strip()
            df.loc[i] = row

        # save the dataframe to Excel
        # get the file path from the entry widget
        file_path = self.text_location_entry.get()

        # split the file path into its components
        dir_name = os.path.dirname(file_path)
        base_name, ext = os.path.splitext(os.path.basename(file_path))

        # modify the file name
        new_base_name = base_name + "1"

        # join the components back together
        new_file_path = os.path.join(dir_name, new_base_name + ext)
        df.to_excel(new_file_path, index=False)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("ERT Log Parser")
    window = MainWindow(root)
    root.mainloop()
