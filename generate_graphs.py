from openpyxl import load_workbook
import matplotlib.pyplot as plt
from graph_options import *
import numpy as np

workbook = load_workbook(filename = "data.xlsx", data_only = True)

sheets = ["1", "2", "3", "4", "5", "6", "7", "8"]

for sheet in sheets:
	worksheet = workbook[sheet]
	
	# Overall number concentration
	figure_name = "overall_number_" + sheet + ".png"
	graph_title = worksheet["A1"].value
	
	x_label = worksheet["A2"].value
	cells = worksheet["A3:A32"]
	x_axis = []
	for cell in cells:
		x_axis.append(cell[0].value)
	
	y_label = worksheet["B2"].value
	cells = worksheet["B3:B32"]
	y_axis = []
	for cell in cells:
		y_axis.append(cell[0].value)
	
	plt.rcParams["figure.figsize"] = 15,15
	plt.tick_params(axis = "both", which = "both", labelsize = label_font - 4)
	plt.xlabel(x_label, fontsize = label_font)
	plt.ylabel(y_label, fontsize = label_font)
	plt.axis([0, 31, 3000, 6000])
	plt.plot(x_axis, y_axis, "ko--", markersize = marker_size)
	plt.title(graph_title, fontsize = label_font)
	plt.savefig(figure_name)
	plt.clf()
	
	
	# Particle size distribution
	figure_name = "size_distribution_" + sheet + ".png"
	graph_title = worksheet["C1"].value
	positions = np.arange(41)
	
	x_label = worksheet["C2"].value
	cells = worksheet["C3:C43"]
	labels = []
	for cell in cells:
		labels.append("%.3f" % cell[0].value)
	
	y_label = worksheet["D2"].value
	cells = worksheet["D3:D43"]
	hist_data = []
	for cell in cells:
		hist_data.append(cell[0].value)
	
	plt.rcParams["figure.figsize"] = 15,15
	plt.tick_params(axis = "both", which = "both", labelsize = label_font - 4)
	plt.xticks(positions, labels, rotation = 70, fontsize = label_font - 12)
	plt.xlabel(x_label, fontsize = label_font)
	plt.ylabel(y_label, fontsize = label_font)
	plt.axis([-1, 41, 0.0001, 10000])
	plt.gca().set_yscale("log")
	plt.bar(positions, hist_data, width = 0.8, color = "grey")
	plt.title(graph_title, fontsize = label_font)
	plt.savefig(figure_name)
	plt.clf()
	
	
	# Mass concentration
	figure_name = "mass_concentration_" + sheet + ".png"
	graph_title = worksheet["E1"].value
	
	x_label = worksheet["E2"].value
	y_label = worksheet["F2"].value
	cells = worksheet["E4:E34"]
	x_axis = []
	for cell in cells:
		x_axis.append(cell[0].value)
	
	y1_label = worksheet["F3"].value
	cells = worksheet["F4:F34"]
	y1_axis = []
	for cell in cells:
		y1_axis.append(cell[0].value)
	
	y2_label = worksheet["G3"].value
	cells = worksheet["G4:G34"]
	y2_axis = []
	for cell in cells:
		y2_axis.append(cell[0].value)
	
	plt.rcParams["figure.figsize"] = 15,15
	plt.tick_params(axis = "both", which = "both", labelsize = label_font - 4)
	plt.xlabel(x_label, fontsize = label_font)
	plt.ylabel(y_label, fontsize = label_font)
	plt.axis([0, 31, 0, 7])
	plt.plot(x_axis, y1_axis, "ko--", markersize = marker_size, label = y1_label)
	plt.plot(x_axis, y2_axis, "r^--", markersize = marker_size, label = y2_label)
	plt.title(graph_title, fontsize = label_font)
	plt.legend(loc = "upper right", fontsize = label_font - 4)
	plt.savefig(figure_name)
	plt.clf()
