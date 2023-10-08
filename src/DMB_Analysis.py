import matplotlib
from skimage.transform import rescale

matplotlib.use('TkAgg')

import tkinter as tk
from collections import OrderedDict
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from tkinter.filedialog import askopenfilename, askopenfilenames
from skimage import io, util, exposure
from skimage.color import rgb2grey, grey2rgb
import numpy as np
from PIL import Image, ImageDraw
from savitsky_golay import savitsky_golay
from xlsxwriter import Workbook
from copy import deepcopy
from os.path import basename, dirname, join, exists, splitext
import re
from skimage.exposure import equalize_adapthist
from skimage.morphology import dilation, disk
import xlrd


RESOLUTION = '720x540'  # Original Resolution of the window
THICKNESS = {50: (0.378, 0.755, 1.133, 1.511, 1.888, 2.266, 2.644, 3.021),
             100: (0.189, 0.378, 0.567, 0.755, 0.944, 1.133, 1.322, 1.511)}  # key: standard thickness;
                                                                             # value: list of the 8 steps aluminium
                                                                             # standard corresponding DMB.

DEFAULT_DMB_THRESHOLD = 0.6
COLOR_DEPTH = 2**16
STD_THRESHOLD = 4096
STD_FILE_NAME = 'Standard.xlsx'
DMB_FILE_NAME = "DMB_Results.xlsx"
XMARGIN = 400
YMARGIN = 350
MAX_SIZE = 2000
CLIP_LIMIT = 0.02



class MainProgram(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.geometry(RESOLUTION)  # Set geometry of window

        # Properly quit the application when clicking the cross icon
        self.protocol('WM_DELETE_WINDOW', self._quit)


        #-------------------------#
        # MEMBERS DECLARATION
        #-------------------------#
        self.fileName = tk.StringVar(value='')
        self.results = []
        self.roiPointsList = []
        self.img_orig = []
        self.img_mat = np.array([])
        self.scale = 1
        self.stdFunc = None
        self.analysed = False
        self.invert = tk.IntVar()


        #-------------------------#
        # MENU BAR CREATION
        #-------------------------#
        menubar = tk.Menu(self)

        # Menu container
        mainMenus = OrderedDict()

        # File menu
        mainMenus['File'] = OrderedDict()
        mainMenus['File']['Load file'] = self.loadFile

        # Edit menu
        mainMenus['Analysis'] = OrderedDict()

        for menu in mainMenus:
            curMenu = tk.Menu(menubar, tearoff=0)
            for subMenu in mainMenus[menu]:
                curMenu.add_command(label=subMenu, command=mainMenus[menu][subMenu])
            menubar.add_cascade(label=menu, menu=curMenu)

        # Display the menu
        self.config(menu=menubar)


        #-------------------------#
        # LEFT TOOL PANEL CREATION
        #-------------------------#

        # Main frame
        self.panelToolFrame = tk.Frame(self)
        self.panelToolFrame.pack(side=tk.LEFT, fill=tk.BOTH)
        nrow = 0

        # Standard thickness selection (radio buttons)
        self.default_thickness = tk.DoubleVar(value=list(THICKNESS.keys())[-1])
        for i, thickness in enumerate(THICKNESS):
            tk.Radiobutton(master=self.panelToolFrame, text="%s µm"%thickness, variable=self.default_thickness, value=thickness).grid(row=nrow, column=i)
        nrow += 1


        self.std_equation = tk.StringVar(value="No equation")
        tk.Label(self.panelToolFrame, textvariable=self.std_equation).grid(row=nrow, column=1)
        self.loadFileStdBut = tk.Button(self.panelToolFrame, text='Load Standard Files', command=self.loadFileStd, bg='red')
        self.loadFileStdBut.grid(row=nrow, column=0)
        nrow += 1

        self.fileNameTxt = tk.Entry(self.panelToolFrame, textvariable=self.fileName, state='disabled')
        self.fileNameTxt.grid(row=nrow, column=0)
        loadFileBut = tk.Button(self.panelToolFrame, text='Load File', command=self.loadFile)
        loadFileBut.grid(row=nrow, column=1)
        nrow += 1

        tk.Label(self.panelToolFrame, text="Field").grid(row=nrow, column=0)
        self.field = tk.StringVar(self.panelToolFrame)
        self.field.set('gen')
        fieldEdit = tk.Entry(self.panelToolFrame, textvariable=self.field)
        fieldEdit.grid(row=nrow, column=1)
        nrow += 1

        tk.Label(self.panelToolFrame, text="Threshold").grid(row=nrow, column=0)
        self.threshold = tk.DoubleVar(self.panelToolFrame)
        self.threshold.set(DEFAULT_DMB_THRESHOLD)
        thresholdEdit = tk.Entry(self.panelToolFrame, textvariable=self.threshold)
        thresholdEdit.grid(row=nrow, column=1)
        nrow += 1

        tk.Label(self.panelToolFrame, text="Thickness").grid(row=nrow, column=0)
        self.thickness = tk.DoubleVar(self.panelToolFrame)
        self.thickness.set(self.default_thickness.get())
        thicknessEdit = tk.Entry(self.panelToolFrame, textvariable=self.thickness)
        thicknessEdit.grid(row=nrow, column=1)
        nrow += 1

        self.invert = tk.IntVar()
        tk.Checkbutton(self.panelToolFrame, text="Invert", variable=self.invert, command=self.invertImg).grid(row=nrow, column=0)
        nrow += 1

        self.contrast_Scale = tk.Scale(self.panelToolFrame, from_=0.01, to=0.1, resolution=0.01, orient="horizontal")
        self.contrast_Scale.grid(row=nrow, column=0, sticky=tk.W + tk.E, columnspan=2)
        self.contrast_Scale.set(CLIP_LIMIT)
        self.contrast_Scale.bind("<ButtonRelease>", self.change_contrast)
        nrow += 1

        tk.Label(self.panelToolFrame, text="Contrast scale").grid(row=nrow, column=0, sticky=tk.W + tk.E, columnspan=2)
        nrow += 1


        # CANVAS PLOT CREATION
        self.fig = plt.figure(figsize=(1, 1))
        self.ax = self.fig.add_subplot(111)
        plotFrame = tk.Frame(self)
        plotFrame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=1)
        self.canvas = FigureCanvasTkAgg(self.fig, master=plotFrame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.toolbar = NavigationToolbar2Tk(self.canvas, plotFrame)
        self.toolbar.update()

        # Standard curve plot
        self.fig_std = plt.figure(figsize=(1, 1))
        self.ax_std = self.fig_std.add_subplot(111)
        plotFrame = tk.Frame(self.panelToolFrame)
        plotFrame.grid(row=nrow, column=0, rowspan=1, columnspan=2, sticky=tk.N + tk.S + tk.W + tk.E, pady=2)
        self.panelToolFrame.rowconfigure(nrow, weight=1)
        nrow += 1
        self.canvas_std = FigureCanvasTkAgg(self.fig_std, master=plotFrame)
        self.canvas_std.draw()
        self.canvas_std.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.ax_std.set_title("Standard peaks", fontsize=8)
        self.ax_std.tick_params(labelsize=6)
        self.ax_std.ticklabel_format(style='sci', scilimits=(0,0), axis='both')
        self.canvas_std.draw()

        # DMB curve plot
        self.fig_dmb = plt.figure(figsize=(1, 1))
        self.ax_dmb = self.fig_dmb.add_subplot(111)
        self.ax_dmb.set_title("ROI histogram", fontsize=8)
        self.ax_dmb.tick_params(labelsize=6)
        self.ax_dmb.ticklabel_format(style='sci', scilimits=(0,0), axis='both')
        plotFrame = tk.Frame(self.panelToolFrame)
        plotFrame.grid(row=nrow, column=0, rowspan=1, columnspan=2, sticky=tk.N + tk.S + tk.W + tk.E, pady=2)
        self.panelToolFrame.rowconfigure(nrow, weight=1)
        nrow += 1
        self.canvas_dmb = FigureCanvasTkAgg(self.fig_dmb, master=plotFrame)
        self.canvas_dmb.draw()
        self.canvas_dmb.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        self.canvas_dmb.draw()

        # Write results button
        tk.Button(self.panelToolFrame, text='Write results', command=self.write_results).grid(row=nrow, column=0)
        tk.Button(self.panelToolFrame, text='Clear results (new sample analysis)', command=self.clear_results).grid(row=nrow, column=1)
        nrow += 1

        # Connecting the matplotlib events
        self.eventId = []
        self.eventId.append(self.canvas.mpl_connect('button_press_event', self.on_click))

        self.bind("<Escape>", self.reset_roi)
        self.bind("<Delete>", self.del_last_point)

        return

    def loadFile(self):
        """ Load sample image file and draw it in the main ax"""
        filename = askopenfilename()
        if not filename:
            return
        self.fileName.set(filename)
        self.fileNameTxt.xview_moveto(1)
        self.img_mat = rgb2grey(io.imread(self.fileName.get()))  # Careful ! Image is now in float (grey levels from 0
                                                                 # to 1)
        self.img_orig = deepcopy(self.img_mat)

        # CONTRAST ENHANCEMENT
        self.img_mat = equalize_adapthist(self.img_mat, clip_limit=CLIP_LIMIT)

        self.drawImage(self.img_mat)
        self.canvas.draw()

        self.roiPointsList = []
        self.contrast_Scale.set(CLIP_LIMIT)
        self.invert.set(0)
        return

    def loadFileStd(self):
        """ Load the standard files and extract the logarithm function to convert gray level to DMB. """
        imPaths = askopenfilenames()
        if not imPaths:
            return

        if imPaths[0].endswith(".xlsx"):
            wb = xlrd.open_workbook(imPaths[0])
            sheet = wb.sheet_by_index(0)
            for row in range(sheet.nrows):
                if sheet.cell_value(row, 0) == "a":
                    a = float(sheet.cell_value(row, 1))
                if sheet.cell_value(row, 0) == "b":
                    b = float(sheet.cell_value(row, 1))
                if sheet.cell_value(row, 0) == "r":
                    r = float(sheet.cell_value(row, 1))

            # Creating function of convertion Grey Level to DMB
            def f(x):
                x = np.array(x)
                return a * np.log10(x) + b

            self.stdFunc = f
            self.std_equation.set('%.2f*GL + %.2f (r=%.3f)' % (a, b, r))

            self.loadFileStdBut.configure(bg="green")
            return
        valCalib = THICKNESS[self.default_thickness.get()]  # Calibration values for each aluminium step
        resFilePath = join(dirname(imPaths[0]), STD_FILE_NAME)  # Path of the standard results file

        n = len(valCalib)  # Number of aluminium steps
        greyScale = np.zeros(STD_THRESHOLD)  # Grey scale histogram of all standard images from 0 to the grey threshold
                                             # of the standard
        x = range(STD_THRESHOLD)  # Grey levels from 0 to the grey threshold of the standard
        for imPath in imPaths:
            imarray = np.array(Image.open(imPath))
            imarray = imarray[YMARGIN:-YMARGIN]
            greyScale += np.histogram(imarray.ravel(), bins=range(0, COLOR_DEPTH+1))[0][:STD_THRESHOLD]

        smoothed = savitsky_golay(greyScale, 31, 3)  # Smoothed histogram with Savitsky-Golay method

        # Detection of the maxima in the histogram
        maxima = []
        for i in range(1, len(smoothed) - 1):
            if smoothed[i] >= smoothed[i - 1] and smoothed[i] > smoothed[i + 1]:
                maxima.append(i+1)

        maxima.sort(key=lambda x: -smoothed[x])  # Maximum index sorted by the total number of pixel it represents
        if len(maxima) < 8:
            print('ERROR, not enough peak detected.')
        realMax = maxima[:8]  # Takes the 8 first maxima (supposedly corresponding to the 8 steps)

        # Plotting standard histogram and maxima
        self.ax_std.clear()
        self.ax_std.plot(smoothed)
        self.ax_std.plot(realMax, smoothed[realMax], 'ro')
        self.ax_std.set_title("Standard peaks", fontsize=8)
        self.ax_std.ticklabel_format(style='sci', scilimits=(0,0), axis='both')
        self.ax_std.tick_params(labelsize=6)
        self.canvas_std.draw()

        # Plotting in colorful pixels the pixels corresponding to each maxima found (for verification purpose)
        self.fig.clear()
        for i, imPath in enumerate(imPaths):
            imarray = np.array(Image.open(imPath))
            imarray = imarray[YMARGIN:-YMARGIN]
            imarrayrgb = grey2rgb(equalize_adapthist(imarray))
            colors = ([1, 0, 0], [0, 1, 0], [0, 0, 1], [1, 1, 0], [1, 0, 1], [0, 1, 1], [1, 0.5, 0.5], [0.5, 1, 0.5])
            for j, max in enumerate(sorted(realMax)):
                mask = imarray == max
                mask = dilation(mask, selem=disk(3))
                imarrayrgb[mask] = colors[j%len(colors)]
            ax = self.fig.add_subplot(len(imPaths), 1, i+1)
            ax.imshow(imarrayrgb)
            self.canvas.draw()


        gLv = sorted(np.array(x)[realMax],
                     key=lambda x: -x)  # Sorting the maxima from the lighter grey to the darker grey

        # Calculating the a and b coefficients for the grey level to dmb conversion function
        X = np.array(gLv)
        Y = np.array(valCalib)
        a = (np.sum(np.log10(X) * Y) - ((np.sum(np.log10(X)) * np.sum(Y)) / n)) / (
                np.sum((np.log10(X)) ** 2) - ((np.sum(np.log10(X))) ** 2 / n))
        b = (np.sum(Y) / n) - a * (np.sum(np.log10(X)) / n)

        # Calculating the coefficient of correlation (for precision assessment)
        sX = np.sqrt(np.sum(1. / n * (np.mean(np.log10(X)) - np.log10(X)) ** 2))
        sY = np.sqrt(np.sum(1. / n * (np.mean(Y) - Y) ** 2))
        covarianceXY = np.mean(np.log10(X) * Y) - (np.mean(np.log10(X)) * np.mean(Y))
        r = covarianceXY / (sX * sY)

        print('a = ', a)
        print('b = ', b)
        print('r = ', r)

        self.std_equation.set('%.2f*GL + %.2f (r=%.3f)'%(a, b, r))

        # Saving the results in an xlsx file
        fileSave = Workbook(resFilePath)
        worksheet = fileSave.add_worksheet()

        col1 = ['Nom', 'Marche 1', 'Marche 2', 'Marche 3', 'Marche 4', 'Marche 5', 'Marche 6', 'Marche 7', 'Marche 8',
                '', 'a', 'b', 'r']
        col2 = ['Niveau de gris'] + list(gLv) + ['', str(a), str(b), str(r)]
        for i in range(len(col1)):
            worksheet.write(i, 0, col1[i])
            worksheet.write(i, 1, col2[i])

        fileSave.close()

        # Creating function of convertion Grey Level to DMB
        def f(x):
            x = np.array(x)
            return a * np.log10(x) + b

        self.stdFunc = f


    def drawImage(self, img_mat):
        """ Clear the main figure (Image display) and draw the new image in it, with a potential rescale of the
        displayed image (for performance purpose)"""
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.ax.set_adjustable('box')
        self.scale = MAX_SIZE/max(self.img_mat.shape)
        if self.scale < 1:
            img_mat = rescale(img_mat, self.scale)
        else:
            self.scale = 1
        self.ax.imshow(img_mat, cmap='gray')

        # Drawing margins of images overlap
        self.ax.axhline(y=YMARGIN*self.scale)
        self.ax.axhline(y=len(img_mat) - YMARGIN*self.scale)
        self.ax.axvline(x=XMARGIN*self.scale)
        self.ax.axvline(x=len(img_mat[0]) - XMARGIN*self.scale)

        self.fig.tight_layout()
        self.toolbar.update()
        self.img_reduced = img_mat

        return

    def invertImg(self):
        try:
            self.img_mat = util.invert(self.img_mat)
            self.drawImage(self.img_mat)
            self.canvas.draw()
        except:
            print("ERROR INVERT IMAGE")

    def change_contrast(self, *args):
        try:
            self.img_mat = exposure.equalize_adapthist(self.img_orig, clip_limit=self.contrast_Scale.get())
            if self.invert.get():
                self.img_mat = util.invert(self.img_mat)
            self.drawImage(self.img_mat)
            self.canvas.draw()
        except:
            print("ERROR CHANGE CONTRAST")

    def on_click(self, event):
        """ Manage the mouse clicking event """
        if self.toolbar.mode:
            return
        if event.button == 1:  # Left button
            # Create a point of the ROI
            if self.analysed:
                self.reset_roi()
                self.analysed = False
            self.roiPointsList.append((event.xdata*1./self.scale, event.ydata*1./self.scale))
            self.updateROI()

        if event.button == 3:  # Right button
            # Close the ROI polygon and analyse it
            if self.roiPointsList[0] != self.roiPointsList[-1]:
                self.roiPointsList.append(self.roiPointsList[0])
            self.updateROI()
            self.analysed = True
            roiPointListReduced = [(int(roiPoint[0]*self.scale), int(roiPoint[1]*self.scale)) for roiPoint in self.roiPointsList]
            img = Image.new('L', (len(self.img_reduced[0]), len(self.img_reduced)), 0)
            ImageDraw.Draw(img).polygon(roiPointListReduced[:-1], outline=1, fill=1)
            mask = np.array(img).astype(float)
            mask[mask == 0] = 0.2
            self.img_roi = self.img_reduced*mask
            img = Image.new('L', (len(self.img_orig[0]), len(self.img_orig)), 0)
            ImageDraw.Draw(img).polygon(self.roiPointsList[:-1], outline=1, fill=1)
            mask = np.array(img)
            self.img_orig_roi = self.img_orig*mask
            self.ax.clear()
            self.ax.imshow(self.img_roi, cmap='gray')
            self.updateROI()
            self.canvas.draw()

            self.analyze_frame()

    def updateROI(self):
        """ Plots the ROI """
        if len(self.roiPointsList) == 1 or self.roiPointsList[0] != self.roiPointsList[-1]:
            style = 'ro-'
        else:
            style = 'r-'
        self.ax.plot([point[0]*self.scale for point in self.roiPointsList],
                     [point[1]*self.scale for point in self.roiPointsList], style, linewidth=0.3, markersize=2)
        self.canvas.draw()

    def reset_roi(self, *arg, **kwargs):
        """ Reset the ROI """
        self.roiPointsList = []
        self.ax.clear()
        self.ax.imshow(self.img_reduced, cmap='gray')
        self.ax.axhline(y=YMARGIN*self.scale)
        self.ax.axhline(y=len(self.img_reduced) - YMARGIN*self.scale)
        self.ax.axvline(x=XMARGIN*self.scale)
        self.ax.axvline(x=len(self.img_reduced[0]) - XMARGIN*self.scale)

        self.canvas.draw()

    def del_last_point(self, *args, **kwargs):
        if self.roiPointsList:
            self.roiPointsList.pop()
            self.ax.clear()
            self.ax.imshow(self.img_reduced, cmap='gray')
            self.ax.axhline(y=YMARGIN*self.scale)
            self.ax.axhline(y=len(self.img_reduced) - YMARGIN*self.scale)
            self.ax.axvline(x=XMARGIN*self.scale)
            self.ax.axvline(x=len(self.img_reduced[0]) - XMARGIN*self.scale)
            self.updateROI()
            self.canvas.draw()

    def analyze_frame(self):
        """ Compute the histogram of gray levels of the original image (not rescaled) and analyse it """
        if self.stdFunc is None:
            print("No gray to DMB conversion function found. Load the aluminium standard files first.")
            return
        bins, dmb_hist = np.unique(self.img_orig_roi.ravel(), return_counts=True)
        bins = self.stdFunc(bins) * self.default_thickness.get() / self.thickness.get()
        arg_thresh = np.argmin(np.abs(bins - self.threshold.get()))
        self.ax_dmb.clear()
        self.ax_dmb.plot(bins[:arg_thresh], dmb_hist[:arg_thresh])
        self.ax_dmb.set_title("ROI histogram", fontsize=8)
        self.ax_dmb.tick_params(labelsize=6)
        self.ax_dmb.ticklabel_format(style='sci', scilimits=(0,0), axis='both')
        self.canvas_dmb.draw()
        self.gl2dmb()


        return

    def gl2dmb(self):
        """ Compute the average DMB, DMB maximum frequence, heterogeneity index and standard deviation of the ROI's
        histogram without taking account of values below the threshold DMB level """
        threshold = self.threshold.get()  # DMB threshold (below it, the analysed matter is not bone
        th = self.thickness.get()  # Actual thickness of the bone cut

        xorig = self.stdFunc(range(1, COLOR_DEPTH))  # DMB value of each grey value in range (1, COLOR_DEPTH)
                                                     # (0 is excluded because it gives infinite value when transformed
                                                     # to DMB)
        argsort = np.argsort(xorig)  # Argument to sort DMB levels because stdFunc might change the direction of
                                     # evolution (cf. stdFunc)
        xorig = xorig[argsort]  # Sorted DMB levels

        imarray = self.img_orig_roi
        greyScale = np.histogram(imarray.ravel(), bins=range(COLOR_DEPTH+1))[0]  # Histogram of the ROI

        greyScale = greyScale[1:][argsort]  # Sorted histogram (cf. argsort)
        x = xorig * self.default_thickness.get() / th  # Corrected DMB with actual thickness
        argthresh = np.abs(x - threshold).argmin()

        nbpix = sum(greyScale[argthresh:])  # Total number of pixels
        avgdmb = sum(greyScale[argthresh:] * x[argthresh:]) / nbpix  # Average DMB
        SD = np.sqrt(sum(greyScale[argthresh:] * (x[argthresh:] - avgdmb) ** 2) / nbpix)  # Standard Deviation

        dmbargmax = np.argmax(greyScale[argthresh:])
        dmbfmax = x[argthresh:][dmbargmax]  # Maximum frequency of DMB
        try:
            x1 = np.argmin(abs(greyScale[argthresh:][:dmbargmax] - greyScale[argthresh:][
                dmbargmax] / 2))  # Left part from the maximum
            x2 = np.argmin(abs(greyScale[argthresh:][dmbargmax:] - greyScale[argthresh:][
                dmbargmax] / 2))  # Right part from the maximum
            HI = abs(x[argthresh:][:dmbargmax][x1] - x[argthresh:][dmbargmax:][x2])  # Heterogeneity Index (FWHM)
        except:
            HI = 'Unknown'

        self.ax_dmb.text(0.02, 0.95, 'Avg DMB = %.2f\nSD = %.2f'%(avgdmb, SD), verticalalignment='top', transform=self.ax_dmb.transAxes)
        self.canvas_dmb.draw()

        resrow = [basename(self.fileName.get()), self.field.get(), th, avgdmb, SD, dmbfmax, HI,
                  nbpix]  # List of the ROI's results

        # Saving the ROI's results into self.results
        replace = False  # Checks if the same file and same field has been analyzed.
                         # If yes replace the already existing result with the new one.
        for i in range(len(self.results)):
            if basename(self.fileName.get()) == self.results[i][0] and self.field.get() == self.results[i][1]:
                self.results[i] = resrow
                replace = True
        if not replace:
            self.results.append(resrow)

        return

    def write_results(self):
        """ Write the results in an xlsx file in the last analysed file's directory """
        if not self.results:
            return
        path = join(dirname(self.fileName.get()), DMB_FILE_NAME)
        i = 0
        while exists(path):
            i += 1
            path = "%s_%i%s" % (splitext(path)[0], i, splitext(path)[1])

        # Natural sorting function (same sorting as Windows)
        def natural_sort_key(s, _nsre=re.compile('([0-9]+)')):
            return [int(text) if text.isdigit() else text.lower()
                    for text in re.split(_nsre, s)]

        # Sorting by field fisrt and then by file name
        self.results = sorted(self.results, key=lambda l: (natural_sort_key(l[1]), natural_sort_key(l[0])))
        results = np.array(self.results)

        fields = np.unique(results[:, 1])

        header = (
        "Filename", "Field", "Thickness", "Average DMB", "SD", "DMB freq. max", "Heterogeneity Index", "Pixel count")

        fileSave = Workbook(path)
        worksheet = fileSave.add_worksheet()

        worksheet.write_row(0, 0, header)

        # Writing the average of all analysed ROIs
        worksheet.write(1, 0, "Total Average")
        worksheet.write(1, 1, "Whole")
        for j in range(2, len(results[0]) - 1):
            worksheet.write(1, j, np.sum(results[:, j].astype(float) * results[:, -1].astype(float)) / np.sum(
                results[:, -1].astype(float)))
        worksheet.write(1, len(results[0]) - 1, np.sum(results[:, -1].astype(float)))

        # Writing the average of each analysed field
        n = 2
        for field in fields:
            worksheet.write(n, 0, "Total %s" % field)
            worksheet.write(n, 1, field)
            for j in range(2, len(results[0]) - 1):
                colField = np.array([line for line in results if line[1] == field])
                worksheet.write(n, j, np.sum(colField[:, j].astype(float)*colField[:, -1].astype(float))/np.sum(colField[:, -1].astype(float)))
            worksheet.write(n, len(results[0]) - 1, np.sum(colField[:, -1].astype(float)))
            n += 1
        n += 1

        # Writing the average of fields which share common letters (e.g. cort1 and cort2)
        for id, field in enumerate(fields):
            field_ = field
            while str.isdigit(field_[-1]):
                field_ = field_[:-1]
            nb_field_ = 0
            indices = []
            for id2, field2 in enumerate(fields):
                field2_ = field2
                while str.isdigit(field2_[-1]):
                    field2_ = field2_[:-1]
                if field_ == field2_:
                    nb_field_ += 1
                    indices.append(id2)
            if nb_field_ < 2 or id != min(indices):
                continue
            worksheet.write(n, 0, "Total %s" % field_)
            worksheet.write(n, 1, field_)
            for j in range(2, len(results[0]) - 1):
                print(fields[indices])
                colField = np.array([line for line in results if line[1] in fields[indices]])
                worksheet.write(n, j, np.sum(colField[:, j].astype(float)*colField[:, -1].astype(float))/np.sum(colField[:, -1].astype(float)))
            worksheet.write(n, len(results[0]) - 1, np.sum(colField[:, -1].astype(float)))
            n += 1
        n += 2

        # Writing the each ROI's results
        for i in range(len(results)):
            if i != 0 and results[i][1] != results[i-1][1]:
                n += 1

            for j in range(len(results[i])):
                res = results[i][j]
                if j > 1:
                    res = float(res)
                worksheet.write(i+n, j, res)

        fileSave.close()

        return

    def clear_results(self):
        self.results = []

    def _quit(self, *args, **kwargs):
        """ Properly quit the program """
        self.quit()
        self.destroy()

        return

program = MainProgram()
tk.mainloop()
