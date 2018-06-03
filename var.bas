Attribute VB_Name = "Module1"
Public max_process
Public min_process
Public max_time
Public min_time
Public scalex1
Public scalex2
Public scaley1
Public scaley2
Public divx
Public divy
Public dx
Public dy
Public toggle
Public current
Public voltage
Public avr_value
Public avr_value2
Public avr_value3
Public sum_loop
Public sum_current
Public sum_voltage

Public real_current
Public real_voltage
Public torque
Public power
Public rpm
Public step_load
Public record As Boolean
Public x1
Public y1
Public y12
Public y13

Public value_x
Public value_y
Public value_y2
Public value_y3

Public num_plot
Public save_datax(5, 500)

Public save_datay(5, 500)
Public ref_datax
Public ref_datay
Public num_data
Public load_data
Public save_numdata(5)
Public scale_time
Public k
Public store_maxtime
Public store_mintime
Public change_scale As Boolean

Public color_pv
Public color_mv
Public color_sp


Public value_pgain As Boolean
Public value_igain As Boolean
Public value_dgain As Boolean
Public c_data
