import argparse
import PyTango as tango
import pyexcel
from pyexcel.ext import ods, xls, xlsx

parser = argparse.ArgumentParser(description='Save Tango Controls configuration info to a spreadsheet')
parser.add_argument('file', help="Specify where to save the info.")
parser.add_argument('--device-class', help="Save only properties of devies for selected class.", default='*')
parser.add_argument('--exclude-property', help="Allows to exclude certain properties from export.", action='append')
parser.add_argument('--tree-selection',
                    help="Limit info to the selected device tree part (to name starts with).",
                    default='')

args = parser.parse_args()

db = tango.Database()

class_list = db.get_class_list(args.device_class)

# create class based view
cls_sheet_data = [['Class', 'Device', 'Property', 'Value'], ]
for cls in class_list:
    # first get class property
    cls_prop_list = db.get_class_property_list(cls)

    cls_prop_values = db.get_class_property(cls, list(cls_prop_list))

    for cls_prop in cls_prop_values.keys():
        cls_sheet_data.append([cls, 'class property', cls_prop, '\n'.join(cls_prop_values[cls_prop])])

    # get devices
    devices = db.get_device_exported_for_class(cls)
    print devices

    for dev in list(devices):
        if str(dev).lower().startswith(args.tree_selection.lower()):
            dev_prop_list = db.get_device_property_list(dev, '*')
            print dev_prop_list
            dev_prop_values = db.get_device_property(dev, list(dev_prop_list))
            for dev_prop in dev_prop_values.keys():
                if dev_prop not in args.exclude_property:
                    cls_sheet_data.append([cls, dev, dev_prop, '\n'.join(dev_prop_values[dev_prop])])

print cls_sheet_data

sheet = pyexcel.Sheet(cls_sheet_data, name="Properties by Classes")
sheet.save_as(args.file)








