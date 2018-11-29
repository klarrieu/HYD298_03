from __future__ import division
#!/usr/bin/python
# Filename: cMorphUnits.py
import arcpy, os, logging
from arcpy.sa import *
import cIO as cio


class MU:
    def __init__(self, *args):

        self.path = os.path.dirname(os.path.realpath(__file__)) + "\\"
        self.workbook = cio.WorkbookContainer(self.path + "mu.xlsx", 0)
        self.license_state = ""
        self.logger = logging.getLogger("mu_info.log")

        self.mu_names = []
        self.mu_names_number = {}
        self.thresh_h_lower = {}
        self.thresh_h_upper = {}
        self.thresh_u_lower = {}
        self.thresh_u_upper = {}
        self.get_mu_data()
        self.write_col = "C"

        try:
            # for metric provide args[1] = 1.0
            self.unit_conv = float(args[1])
        except:            # default: U.S. customary (1.0)
            self.unit_conv = 0.3047992

    def calculate_mu(self, h_raster, u_raster):
        # h_raster: STR full path to depth raster
        # u_raster: STR full path to velocity raster

        if not str(self.license_state) == "CheckedOut":
            self.license_state = arcpy.CheckOutExtension('Spatial')  # check out license if not yet done

        __ras__ = []  # list of calculated rasters
        try:
            self.logger.info(" > Reading depth and velocity rasters ...")
            h_ras = arcpy.Raster(h_raster)
            u_ras = arcpy.Raster(u_raster)
        except ValueError:
            self.logger.info("ERROR: Cannot find rasters.")
            return -1

        self.logger.info(" > Calculating morphological unit rasters ...")
        for mu in self.mu_names:
            try:
                self.logger.info("   - type: " + mu + " ...")
                __ras__.append(Con(((Float(h_ras) < (self.thresh_h_upper[mu] / self.unit_conv)) & (
                                            Float(u_ras) < (self.thresh_u_upper[mu]) / self.unit_conv)),
                                   Con(((Float(h_ras) >= (self.thresh_h_lower[mu] / self.unit_conv)) & (
                                                Float(u_ras) >= (self.thresh_u_lower[mu] / self.unit_conv))),
                                                                                        int(self.mu_names_number[mu]))))
            except:
                self.logger.info("ERROR: In " + str(mu) + " (revise thresholds and rasters).")
        self.logger.info("   * OK")
        try:
            self.logger.info(" > Combining morphological unit rasters ...")
            return CellStatistics(__ras__, "SUM", "DATA")
        except:
            self.logger.info("ERROR: Could not calculate CellStatistics (raster comparison).")

    def get_mu_data(self):
        self.logger.info(" > Reading data from mu.xlsx ...")
        self.mu_names = self.workbook.read_column("B", 5)
        chr_no = 67
        thresholds = []
        while True:
            if not self.workbook.test_cell_content(chr(chr_no), 5):
                break  # stop reading when first row of a column is empty (col G)
            self.logger.info("   - " + str(self.workbook.read_cell(chr(chr_no), 3)) + ", " + str(self.workbook.read_cell(chr(chr_no), 4)))
            thresholds.append(self.workbook.read_column(chr(chr_no), 5))
            chr_no += 1

        # zip mu names with threshold values (hard coded indices -> weak solution!)
        self.thresh_h_lower = dict(zip(self.mu_names, thresholds[0]))
        self.thresh_h_upper = dict(zip(self.mu_names, thresholds[1]))
        self.thresh_u_lower = dict(zip(self.mu_names, thresholds[2]))
        self.thresh_u_upper = dict(zip(self.mu_names, thresholds[3]))

        # assign unique numbers to all mus
        count = 1
        for mu in self.mu_names:
            self.mu_names_number.update({mu: count})
            count += 1
        self.reload_mu_workbook(1)
        self.logger.info("   * OK")

    def mu_maker(self, h_raster, u_raster, full_out_ras_name, full_out_shp_name, *mu):
        # h_raster: STR - full path to depth raster
        # u_raster: STR - full path to velocity raster
        # full_out_ras_name: STR - full path of the results raster name
        # full_out_shp_name: STR - full path of the result shapefile name
        # mu = LIST(STR) - (optional) - restricts analysis to a list of morphological units according to mu.xlsx

        # start with raster calculations
        self.logger.info("Raster Processing    --- --- ")
        self.license_state = arcpy.CheckOutExtension('Spatial')  # check out license
        arcpy.gp.overwriteOutput = True
        arcpy.env.workspace = self.path
        arcpy.env.extent = "MAXOF"

        try:
            self.mu_names = mu[0]  # limit mu analysis to optional list, if provided
        except:
            pass

        out_ras = self.calculate_mu(h_raster, u_raster)

        try:
            self.logger.info(" > Saving Raster ...")
            out_ras.save(full_out_ras_name)
            self.logger.info("   * OK")
        except:
            self.logger.info("ERROR: Could not save MU raster.")
        arcpy.CheckInExtension('Spatial')  # release license
        self.logger.info("Raster Processing OK     --- \n")

        self.logger.info("Shapefile Processing --- --- ")
        self.logger.info(" > Converting mu raster to shapefile ...")
        arcpy.RasterToPolygon_conversion(arcpy.Raster(full_out_ras_name), full_out_shp_name.split(".shp")[0] + "1.shp", "NO_SIMPLIFY")
        self.logger.info(" > Calculating Polygon areas ...")
        arcpy.CalculateAreas_stats(full_out_shp_name.split(".shp")[0] + "1.shp", full_out_shp_name)
        self.logger.info("   * OK - Removing remainders ...")
        arcpy.Delete_management(full_out_shp_name.split(".shp")[0] + "1.shp")
        self.logger.info(" > Adding MU field ...")
        arcpy.AddField_management(full_out_shp_name, "MorphUnit", "TEXT", "", "", 50)
        expression = "the_dict[!gridcode!]"
        codeblock = "the_dict = " + str(dict(zip(self.mu_names_number.values(), self.mu_names_number.keys())))
        arcpy.CalculateField_management(full_out_shp_name, "MorphUnit", expression, "PYTHON", codeblock)
        self.logger.info("Shapefile Processing OK  --- ")

    def reload_mu_workbook(self, worksheet):
        # worksheet = INT
        self.workbook.save_close_wb(self.path + "mu.xlsx")
        self.workbook = cio.WorkbookContainer(self.path + "mu.xlsx", worksheet)

    def release_mu_workbook(self):
        self.workbook.save_close_wb(self.path + "mu.xlsx")

    def write_area2wb(self, full_shp_name, discharge):
        # full_out_shp_name: STR - full path of the result shapefile name
        # discharge: STR - required for column name
        # reads areas for each MU in full_shp_name, sums it up and writes area to workbook
        self.logger.info(" > Evaluating MU areas ...")
        self.workbook.write_data_cell(self.write_col, 3, str(discharge))
        mu_count = 4
        for mu in self.mu_names:
            try:
                # write MU name and check consistency between analyzed MU and workbook MU
                if not self.workbook.test_cell_content("B", mu_count):
                    self.workbook.write_data_cell("B", mu_count, mu)
                else:
                    if not(self.workbook.read_cell("B", mu_count) == mu):
                        self.logger.info("WARNING: Analyzed MU differs from workbook (results) MU.")
                        self.logger.info("         - Q = " + str(discharge))
                        self.logger.info("         - MU(cell) = " + str(self.workbook.read_cell("B", mu_count)))
                        self.logger.info("         - MU(current) = " + str(mu))

                area = 0.0
                with arcpy.da.UpdateCursor(full_shp_name, ["MorphUnit", "F_AREA"]) as cursor:
                    for row in cursor:
                        try:
                            if str(row[0]).lower().strip() == str(mu).lower().strip():
                                area += float(row[1])
                        except:
                            print("Bad area value.")
                self.workbook.write_data_cell(self.write_col, mu_count, area)
                self.logger.info("   - Area of " + mu + " = %.1f sqft (written to workbook)." % area)
            except:
                self.logger.info("ERROR: Could not read area for %s." % mu)
            mu_count += 1
        self.write_col = self.workbook.col_increase_letter(self.write_col)
        self.logger.info("   * OK")

    def __call__(self, *args, **kwargs):
        print("Class Info: <type> = cMorphUnits.MU")
