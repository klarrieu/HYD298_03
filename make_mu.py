import os, glob, webbrowser
import fFree as ff
import cMorphUnits as cmu
import cIO as cio


def main(raster_dir, *mu_sublist):
    logger = ff.logging_begin("mu_info.log")

    # read discharges related to depth raster names and pack into discharges dictionary
    discharge_workbook = cio.Read(os.path.dirname(os.path.realpath(__file__)) + "\\discharge_info.xlsx", 0, "mu_info.log")
    discharges = dict(zip(discharge_workbook.read_column("E", 5), discharge_workbook.read_column("B", 5)))
    discharge_workbook.close_wb()

    mu_processor = cmu.MU()
    rasters = dict(zip(glob.glob(raster_dir + "input\\h**.aux.xml"), glob.glob(raster_dir + "input\\u**.aux.xml")))

    for h_ras in rasters.keys():
        q = int(discharges[h_ras.split("\\")[-1].split(".aux")[0]])  # get discharge
        logger.info("\n \nDISCHARGE: %d" % q)

        # assign MU raster and shapefile output names for this discharge
        out_ras_name = raster_dir + "output\\q%006dcfs" % q
        out_shp_name = os.path.dirname(os.path.realpath(__file__)) + "\\geodata\\shapefiles\\" + "q%006dcfs" % q + ".shp"

        # make MU rasters and shapefile and write results to workbook
        h_raster = h_ras.split(".aux")[0]
        u_raster = rasters[h_ras].split(".aux")[0]
        try:
            # limit morphological unit analysis to a list of selected mu, if provided (mu = [[slackw..., pol,...]])
            logger.info("Analysis limited to:\n ** " + "\n ** ".join(mu_sublist[0]))
            mu_processor.mu_maker(h_raster, u_raster, out_ras_name, out_shp_name, mu_sublist[0])
        except:
            mu_processor.mu_maker(h_raster, u_raster, out_ras_name, out_shp_name)

        mu_processor.write_area2wb(out_shp_name, q)

    mu_processor.release_mu_workbook()  # save area results in workbook and closes workbook
    ff.logging_end(logger)
    try:
        # open mu.xlsx
        webbrowser.open(os.path.dirname(os.path.realpath(__file__)) + "\\mu.xlsx")
    except ValueError:
        pass


if __name__ == "__main__":
    main(os.getcwd() + "\\geodata\\rasters\\")
