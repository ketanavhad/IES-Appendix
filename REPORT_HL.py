import iesve
import tkinter as tk
import xlsxwriter
import os

def generate_window(project):

    class Window(tk.Frame):
        def __init__(self, master=None):
            tk.Frame.__init__(self, master)

            self.project = project
            self.project_folder = project.path
            self.save_file_name = 'Room_Data'
            self.master = master
            self.init_window()

        # Creation of init_window
        def init_window(self):

            # changing the title of our master widget
            self.master.title("Room Data")
            self.master.columnconfigure(0, weight=1)
            self.master.rowconfigure(0, weight=1)
            self.master.grid()

            print(self.project_folder)
            tk.Label(self, text=' ').grid(row=7, sticky=tk.W)

            tk.Label(self, text='Room Data will be added to an Excel sheet that will be saved into the main project folder').grid(row=8, sticky=tk.W)
            tk.Label(self, text='Name the Excel file below:').grid(row=9, sticky=tk.W)
            self.save_file_entry_box = tk.Entry(self)
            self.save_file_entry_box.insert(0, self.save_file_name)
            self.save_file_entry_box.grid(row=10, sticky='ew')
            tk.Label(self, text=' ').grid(row=11, sticky=tk.W)

            # creating a button instance
            tk.Button(self, text="Run Calculation", command=self.run_calc).grid(row=12, sticky=tk.W)
            
            self.columnconfigure(0, weight=1)
            self.grid(row=0, column=0, sticky='nsew')

        def run_calc(self):
            self.save_file_name = self.save_file_entry_box.get()
            print('Excel File name = \t\t' + self.save_file_name)

            # create excel workbook
            workbook = xlsxwriter.Workbook(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            # create excel work sheet
            sheet1 = workbook.add_worksheet('sheet1')
            # insert image to worksheet
            sheet1.insert_image('A2', 'HL.png',{'x_scale': 0.1, 'y_scale': 0.1})


            def get_room_data(bodies):

                all_room_data = []
                for body in bodies:
                    # create VERoomData object from VEBody object
                    room_data = body.get_room_data(type=0)
                    general_room_data = room_data.get_general()

                    room_name = general_room_data['name']
                    floor_area = round(general_room_data['floor_area'], 1)
                    volume = round(general_room_data['volume'], 1)
                    thermal_template = general_room_data['thermal_template_name']

                    int_gains = room_data.get_internal_gains()
                    occupancy = 0
                    lighting_power = 0
                    equipment_power = 0
                    

                    for int_gain in int_gains:
                        gains_data = int_gain.get()
                        if isinstance(int_gain, iesve.RoomPeopleGain):
                            occupancy = gains_data['occupancies'][0]
                        elif isinstance(int_gain, iesve.RoomLightingGain):
                            lighting_power = gains_data['max_power_consumptions'][0]
                        elif isinstance(int_gain, iesve.RoomPowerGain):  
                            equipment_power = gains_data['max_power_consumptions'][0]
                        else:
                            print("warning: unknown internal gain type") 

                    air_exchanges = room_data.get_air_exchanges()
                    infiltration = 0
                    nat_vent = 0
                    aux_vent = 0

                    for air_ex in air_exchanges:
                        air_ex_data = air_ex.get()
                        if air_ex_data['type_val'] == 0:
                            infiltration = round(air_ex_data['max_flows'][0],1)
                        if air_ex_data['type_val'] == 1:
                            nat_vent = round(air_ex_data['max_flows'][0],1)
                        if air_ex_data['type_val'] == 2:
                            aux_vent = round(air_ex_data['max_flows'][0],1)

                    room_data = [room_name,
                                 floor_area,
                                 volume,
                                 thermal_template,
                                 occupancy,
                                 lighting_power,
                                 equipment_power,
                                 infiltration,
                                 aux_vent,
                                 nat_vent]

                    all_room_data.append(room_data)

                return all_room_data

            # create a list of VEModel objects
            models = project.models
            # The 'real building' model is always the first model in the 'models' list
            model = models[0]
            print('Building Model = ' + model.id)
            # create list of VEBody objects from VEModel object
            bodies = model.get_bodies(False)

            # run main calculation functions
            print('Running Calculations')
            room_data = get_room_data(bodies)
            heading = ['Room Name',
                       'Floor Area (m2)',
                       'Volume (m3)',
                       'Thermal Template',
                       'Occupancy Density (m2/person)',
                       'Lighting Power (W/m2)',
                       'Equipment Power (W/m2)',
                       'Infiltration (ach)',
                       'Auxiliary Ventilation (ach)',
                       'Natural Ventilation (ach)']

            # write data to excel worksheets
            print('Writing results to Excel Sheet')

            # write results data
            y = 6
            sheet1.write_row(y-1, 0, heading)
            for data in room_data:
                sheet1.write_row(y, 0, data)
                y += 1

            # set column widths
            sheet1.set_column('A:J', 20)
            try:
                workbook.close()
            except PermissionError as e:
                print("Couldn't close workbook: ", e)
            os.startfile(self.project_folder + '\\' + self.save_file_name + '.xlsx')
            root.destroy()

    root = tk.Tk()
    app = Window(root)
    root.mainloop()

if __name__ == '__main__':
    project = iesve.VEProject.get_current_project()

    generate_window(project)