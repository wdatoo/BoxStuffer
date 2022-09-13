from django.shortcuts import render
from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic import View
from tablib import Dataset
from ortools.linear_solver import pywraplp
import xlsxwriter
import io
from django.contrib import messages


class HomeView(View):
    def get(self, request, *args, **kwargs):
        context = {}
        return render(request, "pages/home.html", context=context)


class FileView(View):
    def get(self, request, *args, **kwargs):
        context = {}
        return render(request, "pages/home.html", context=context)

    def post(self, request, *args, **kwargs):
        dataset = Dataset()
        new_resource = request.FILES["file"]
        if not new_resource.name.endswith("xlsx"):
            messages.error(request, "Please upload a XLSX file.")
            return render(
                request,
                "pages/home.html",
            )
        get_data = dataset.load(new_resource.read(), format="xlsx")
        my_dic = {}
        weights = []
        trucks = []
        bundle_numbers = []
        weights_nw = []
        for data in get_data:
            weights.append(data[2])
            trucks.append(data[0])
            bundle_numbers.append(data[1])
            weights_nw.append(data[3])
        my_dic["weights"] = weights
        my_dic["items"] = list(range(len(weights)))
        my_dic["bins"] = my_dic["items"]
        if not request.POST["other_value"]:
            my_dic["bin_capacity"] = float(request.POST["bin_capacity"])
        else:
            my_dic["bin_capacity"] = float(request.POST["other_value"])
        # Create the mip solver with the SCIP backend.
        solver = pywraplp.Solver.CreateSolver("SCIP")
        if not solver:
            return
        # Variables
        # x[i, j] = 1 if item i is packed in bin j.
        x = {}
        for i in my_dic["items"]:
            for j in my_dic["bins"]:
                x[(i, j)] = solver.IntVar(0, 1, "x_%i_%i" % (i, j))
        # y[j] = 1 if bin j is used.
        y = {}
        for j in my_dic["bins"]:
            y[j] = solver.IntVar(0, 1, "y[%i]" % j)
        # Constraints
        # Each item must be in exactly one bin.
        for i in my_dic["items"]:
            solver.Add(sum(x[i, j] for j in my_dic["bins"]) == 1)
        # The amount packed in each bin cannot exceed its capacity.
        for j in my_dic["bins"]:
            solver.Add(
                sum(x[(i, j)] * my_dic["weights"][i] for i in my_dic["items"])
                <= y[j] * my_dic["bin_capacity"]
            )
        # Objective: minimize the number of bins used.
        solver.Minimize(solver.Sum([y[j] for j in my_dic["bins"]]))
        status = solver.Solve()
        my_list = []
        if status == pywraplp.Solver.OPTIMAL:
            num_bins = 0.0
            for j in my_dic["bins"]:
                if y[j].solution_value() == 1:
                    bin_items = []
                    bin_weight = 0
                    for i in my_dic["items"]:
                        if x[i, j].solution_value() > 0:
                            bin_items.append(i)
                            bin_weight += my_dic["weights"][i]
                    if bin_weight > 0:
                        num_bins += 1
                        js = {}
                        js["num"] = j
                        js["items"] = bin_items
                        js["total"] = bin_weight
                        my_list.append(js)
            for i in my_list:
                com_list = []
                for ii in i["items"]:
                    com_dict = {}
                    get_truck = trucks[ii]
                    get_bundle = bundle_numbers[ii]
                    get_nw = weights_nw[ii]
                    get_gw = weights[ii]
                    com_dict["truck"] = get_truck
                    com_dict["bundle"] = get_bundle
                    com_dict["nw"] = get_nw
                    com_dict["gw"] = get_gw
                    com_list.append(com_dict)
                i["data"] = com_list
        else:
            messages.error(request, "The problem does not have an optimal solution.")
            return render(request, "pages/home.html")
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()
        row = 0
        col = 0
        for i in my_list:
            worksheet.write(row, col, "container number: {}".format(i["num"]))
            row = row + 2
            worksheet.write(row, col, "Truck")
            col = col + 1
            worksheet.write(row, col, "bundle")
            col = col + 1
            worksheet.write(row, col, "NW weight")
            col = col + 1
            worksheet.write(row, col, "GW weight")
            col = 0
            row = row + 1
            for ii in i["data"]:
                worksheet.write(row, col, ii["truck"])
                col = col + 1
                worksheet.write(row, col, ii["bundle"])
                col = col + 1
                worksheet.write(row, col, ii["nw"])
                col = col + 1
                worksheet.write(row, col, ii["gw"])
                row = row + 1
                col = 0
            col = col + 2
            worksheet.write(row, col, i["total"])
            col = 0
            row = row + 2
        workbook.close()
        output.seek(0)
        filename = "{}.xlsx".format(request.POST["file_name"])
        response = HttpResponse(
            output,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = "attachment; filename=%s" % filename
        return response
