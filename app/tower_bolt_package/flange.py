# -*- coding: utf-8 -*-
"""
Created on Wed Jan 10 11:11:12 2024

@author: TOBHI
"""
import pandas as pd
from datetime import datetime as dt
from datetime import timedelta
import tower_bolt_package.funcs as funcs
import numpy as np


class Flange:

    def __init__(self, path, location, criteria):
        """
        Object used to represent a given flange on a tower which will be analyzed.

        Attributes
        ----------
        path : str
            String of path to flange folder location on disk.
        location : dict
            Dict of the str names of the project, tower, and flange code of the 
            flange to be analyzed.
        headers : Pandas DataFrame
            Dataframe with the combined header data from the two rounds on the flange.
        records : Pandas DataFrame
            Dataframe with the combined record data from the two rounds on the flange.
        stats : dict
            Dict of two dataframes with the count and average/deviance bolt stats.
        errors : str
            String that is added to when known errors are encountered.
        xml_data : dict
            Dict including the separated round 1 and round 2 data before being
            combined.
        rotation_dict : dict
            Dict listing the required rotation levels for each bolt size.
        response : str
            Unused so far. Will include the determined response to recommend repair 
            actions.

        Methods
        -------
        __get_data(file_round:str)
            Fetches and parses the data from an Xml file into the round data.
        __eval_headers()
            Compares the header data between round 1 and round 2.
        __eval_bolts()
            Evaluates the bolt rotation data and compares to required rotation to 
            determine failed bolts.
        __get_stats()
            Determines the number of bolts rotated per cycle and the mean /
            standard devation of rotation for each round.
        __generate_response()
            Uses the failed bolts and the canned responses to generate a response to 
            recommend repair actions to the site.
        run()
            Combined above function to run full analysis on flange.
        """
        self.path = path
        self.location = {"project": location["project"],
                         "tower":   location["tower"],
                         "flange":  location["flange"]}
        self.headers = None
        self.records = None
        self.stats = {"count": None,
                      "total": None}
        self.errors = ""
        self.required_rotation = 0
        self.has_run = 0
        self.xml_data = {
            "first": {"path": None,
                      "headers": None,
                      "records": None},
            "second": {"path": None,
                       "headers": None,
                       "records": None}}
        self.rotation_dict = {
            "M36": 80,
            "M42": 80,
            "M48": 80,
            "M56": 100,
            "M64": 110,
            "M72": 120, }
        self.response = ""
        self.criteria = criteria

    def __get_data(self, file_round: str):
        """
        Fetches and parses the data from an Xml file into the round data.

        Parameters
        ----------
        file_round : str
            String denoting the round to fetch data for. Either "first" or
            "second"

        Returns
        -------
        data : dict
            dict of the header and record data parsed from the Xml file.

        """
        # Discover matching Xml files
        matches = funcs.discover_xmls(self.path, file_round)
        if len(matches) == 0:
            self.errors += (
                f"No {file_round} round Xml file found.")
        elif len(matches) > 1:
            self.errors += (
                f"Unable to determine correct {file_round} round Xml file in folder.")

        # Assumes there is only one of the Xmls found. This overrides the errors
        # if multiple are found, but the error still shows up.
        xml_path = matches[0]
        # Parse the Xml file
        xml_data = funcs.parse_round(xml_path)
        if not xml_data.empty:
            data = {"path": xml_path,
                    "headers": xml_data["headers"],
                    "records": xml_data["records"]}
            self.xml_data[file_round] = data
            return data
        else:
            self.errors += (
                f"Unable to parse from {file_round} round data.")
            return {}

    def __eval_headers(self):
        """
        Compares the header data between round 1 and round 2.

        Returns
        -------
        headers : Pandas DataFrame
            Combined header data with approvals column.

        """
        # Construct combined headers dataframe
        headers = pd.concat(objs=[self.xml_data["first"]["headers"],
                                  self.xml_data["second"]["headers"]],
                            axis=1,
                            ignore_index=True)
        headers = headers.rename(columns={0: "First Round", 1: "Second Round"})
        indicies = headers.index.to_list()
        headers["Approval"] = "N/A"

        # Iterate through header parameters to evaluate
        i = 0
        for index in indicies:
            val1 = headers['First Round'][index]
            val2 = headers['Second Round'][index]
            # ingore differences in caps
            try:
                val1 = val1.upper()
            except:
                pass
            try:
                val2 = val2.upper()
            except:
                pass

            # Date
            # Date of 2nd must be within 72 hours after date of 1st
            # US Date patterns are tested first
            date_patterns = ["%m/%d/%Y", "%m/%d/%Y %H:%M:%S",
                             "%d/%m/%Y", "%d/%m/%Y %H:%M:%S", ]
            if index == "Date":
                for pattern in date_patterns:
                    try:
                        val1 = dt.strptime(val1, pattern)
                        val2 = dt.strptime(val2, pattern)
                        break
                    except:
                        pass
                #  Match date times and make sure the second round is after the first round or at least on the same day
                if (isinstance(val1, str) or isinstance(val2, str)):
                    headers['Approval'][index] = "Alert"
                    self.errors += "Unable to detect and compare dates."
                elif (val2-val1) > timedelta(hours=72):  # 2nd happened more than 72 hrs after 1st
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nSecond round more than 72 hours after first round."
                elif val1 <= val2:  # 2nd happened before 1st
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nDate: Second round may have ocurred before first round."

            # Software Version
            # Software version should match eachother
            elif index == "SoftwareVersion":
                if val1 == val2:
                    # Set up to allow for later checking against reference updateed software version
                    if True:
                        headers['Approval'][index] = "Pass"
                    else:
                        headers['Approval'][index] = "Alert"
                        current_version = ""
                        self.errors += "\nSoftware version differs most recent listed version: "+current_version
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nSoftware version differs between 1st and 2nd round Xml files."

            # Program ID
            # ProgramID must include "first" in 1st and "second" in 2nd
            elif index == "ProgramID":
                if 'first' in val1.lower() and 'second' in val2.lower():
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nProgram Id's do not match expected values."
                pass

            # Bolt Type
            # Not used
            elif index == "BoltType":
                pass

            # Turbine VUI
            # Not used
            elif index == "TurbineVUI":
                pass

            # Tower VUI
            # TowerVUI should match between 1st and 2nd
            elif index == "TowerVUI":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nTower VUI differs between 1st and 2nd round Xml files."

            # Bolt VUI
            # BoltVUI must match between 1st and 2nd
            elif index == "BoltVUI":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nBolt VUI differs between 1st and 2nd round Xml files."

            # Tensioner VUI
            # TensionerVUI should match between 1st and 2nd
            elif index == "TensionerVUI":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nTensioner VUI differs between 1st and 2nd round Xml files."

            # Pump VUI
            # PumpVUI should match between 1st and 2nd
            elif index == "PumpVUI":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nPump VUI differs between 1st and 2nd round Xml files."

            # Operator ID
            # OperatorID should match between 1st and 2nd
            elif index == "OperatorID":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nOperator ID differs between 1st and 2nd round Xml files."

            # Operator Name
            # OperatorName should match between 1st and 2nd
            elif index == "OperatorName":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nOperator Name differs between 1st and 2nd round Xml files."

            # Company
            # Company should match between 1st and 2nd
            elif index == "Company":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Alert"
                    self.errors += "\nCompany differs between 1st and 2nd round Xml files."

            # Bolt Size
            # BoltSize must match between 1st and 2nd
            elif index == "BoltSize":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                    # Set required rotation when they match
                    self.required_rotation = self.rotation_dict[val1]
                else:
                    headers['Approval'][index] = "Fail"
                    # Unable to determine the bolt size for required rotation
                    self.required_rotation = 0
                    self.errors += "\nBolt size differs between 1st and 2nd round Xml files."

            # Bolt QTY
            # BoltQTY must match between 1st and 2nd
            elif index == "BoltQTY":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nBolt QTY differs between 1st and 2nd round Xml files."

            # Clamping Length
            # Not used
            elif index == "ClampingLength":
                pass

            # Flange Location
            # FlangeLocation must match between 1st and 2nd
            elif index == "FlangeLocation":
                if val1 == val2:
                    headers['Approval'][index] = "Pass"
                else:
                    headers['Approval'][index] = "Fail"
                    self.errors += "\nFlange Location differs between 1st and 2nd round Xml files."

            # Angle Sensor Reset Force
            # Not used
            elif index == "AngleSensorResetForce":
                pass

            # Minimum Bolt Tensioning Pressure
            # Not used
            elif index == "MinBoltTensioningPressure":
                pass

            # Minimum Bolt Tensioning Force
            # Not used
            elif index == "MinBoltTensioningForce":
                pass

            # Minimum Nut Rotation Angle First
            # Not used
            elif index == "MinNutRotationAngleFirst":
                pass

            # Maximum Nut Rotation Angle Last
            # Not used
            elif index == "MaxNutRotationAngleLast":
                pass

            # Minimum Nut Loosening Angle
            # Not used
            elif index == "MinNutLooseningAngle":
                pass

            # Minimum Nut Torque
            # Not used
            elif index == "MinNutTorque":
                pass

            # Initial Mean Settlement
            # Not used
            elif index == "InitialMeanSettlement":
                pass

            # Initial Max Settlement
            # Not used
            elif index == "InitialMaxSettlement":
                pass

            # Minimum Required Mean Clamping Force
            # Not used
            elif index == "MinRequiredMeanClampingforce":
                pass

            # No First Tensioning Process
            # Not used
            elif index == "NoFirstTensioningProcess":
                pass

            # No Last Tensioning Process
            # Not used
            elif index == "NoLastTensioningProcess":
                pass

            # Tightenings QTY
            # Not used
            elif index == "TighteningsQTY":
                pass

            # Mean Residual Force
            # Not used
            elif index == "MeanResidualForce":
                pass

            # Flange Approval Tightentings QTY
            # Not used
            elif index == "FlangeApprovalTighteningsQTY":
                pass

            # Flange Approval First Tightening
            # Not used
            elif index == "FlangeApprovalFirstTightening":
                pass

            # Flange Approval Last Tightening
            # Not used
            elif index == "FlangeApprovalLastTightening":
                pass

            # Flange Approval Mean Residual Force
            # Not used
            elif index == "FlangeApprovalMeanResidualForce":
                pass

            i += 1
        self.headers = headers
        return headers

    def __eval_bolts(self):
        """
        Evaluates the bolt rotation data and compares to required rotation to 
        determine failed bolts.

        Returns
        -------
        records : Pandas DataFrame
            Combined important record data for bolts from both rounds with
            approval/failure flags

        """
        # First Round
        # Get relevant data from round
        records1 = self.xml_data["first"]["records"]
        cols = [
            col for col in records1.columns if "BoltRotationAngle" in col and "Target" not in col]
        cols.insert(0, "BoltNo")
        cols.insert(-1, "Cycles")
        records1 = records1[cols]
        # Sum the bolt rotation angles of any cycles #3 or higher
        cycle_sums = np.zeros(len(records1["BoltRotationAngleCycle3"]))
        for col in records1.columns:
            if "BoltRotationAngleCycle" in col:
                try:
                    N = int(col[22:])
                    if N >= 3:
                        cycle_sums += records1[col].to_numpy()
                        records1 = records1.drop(col, axis='columns')
                except Exception as e:
                    print(e)
                    pass
        # Reinsert summed rotations at position
        records1.insert(4, 'BoltRotationAngleCycle3', cycle_sums)

        # Secind Round
        # Get relevant data from round
        records2 = self.xml_data["second"]["records"]
        cols = [
            col for col in records2.columns if "BoltRotationAngle" in col and "Target" not in col]
        cols.insert(0, "BoltNo")
        cols.insert(-1, "Cycles")
        records2 = records2[cols]
        # Sum the bolt rotation angles of any cycles #3 or higher
        cycle_sums = np.zeros(len(records2["BoltRotationAngleCycle3"]))
        for col in records2.columns:
            if "BoltRotationAngleCycle" in col:
                try:
                    N = int(col[22:])
                    if N >= 3:
                        cycle_sums += records2[col].to_numpy()
                        records2 = records2.drop(col, axis='columns')
                except Exception as e:
                    print(e)
                    pass
        # Reinsert summed rotations at position
        records2.insert(4, 'BoltRotationAngleCycle3', cycle_sums)

        # Combine into total records
        records = records1.merge(records2, how='outer', on='BoltNo',)
        # Multi-index to organize
        columns = [("BoltNo", ""),
                   ("First Round", "BoltRotationAngle_x"),
                   ("First Round", "BoltRotationAngleCycle1_x"),
                   ("First Round", "BoltRotationAngleCycle2_x"),
                   ("First Round", "BoltRotationAngleCycle3_x"),
                   ("First Round", "Cycles_x"),
                   ("Second Round", "BoltRotationAngle_x"),
                   ("Second Round", "BoltRotationAngleCycle1_y"),
                   ("Second Round", "BoltRotationAngleCycle2_y"),
                   ("Second Round", "BoltRotationAngleCycle3_y"),
                   ("Second Round", "Cycles_y")
                   ]
        records.columns = pd.MultiIndex.from_tuples(columns)
        records = records.rename(columns={
            'BoltRotationAngle_x': 'Round Total',
            'BoltRotationAngleCycle1_x': 'Cycle 1',
            'BoltRotationAngleCycle2_x': 'Cycle 2',
            'BoltRotationAngleCycle3_x': 'Cycle 3+',
            'Cycles_x': '# Cycles',
            'BoltRotationAngle_y': 'Round Total',
            'BoltRotationAngleCycle1_y': 'Cycle 1',
            'BoltRotationAngleCycle2_y': 'Cycle 2',
            'BoltRotationAngleCycle3_y': 'Cycle 3+',
            'Cycles_y': '# Cycles',
        })

        # Sum all cycles to compare to listed total round rotations
        rd1_total = (records["First Round"]["Cycle 1"] +
                     records["First Round"]["Cycle 2"] +
                     records["First Round"]["Cycle 3+"])
        rd2_total = (records["Second Round"]["Cycle 1"] +
                     records["Second Round"]["Cycle 2"] +
                     records["Second Round"]["Cycle 3+"])
        # Allow a buffer to compare to total rotations for rounding errors
        buffer = self.criteria["Values"]["sum_rounding_buffer"]
        if ((rd1_total - records["First Round"]["Round Total"]).abs().max()) > buffer:
            self.errors += "\nBolt rotation total in round 1 Xml does not match cycles."
        if ((rd2_total - records["Second Round"]["Round Total"]).abs().max()) > buffer:
            self.errors += "\nBolt rotation total in round 2 Xml does not match cycles."

        records = records[records.columns.drop(
            list(records.filter(regex="BoltRotationAngle")))]

        records['Total Rotation'] = (records["First Round"]["Round Total"].fillna(0) +
                                     records["Second Round"]["Round Total"].fillna(0))

        # If we have a determined required rotation, calculate the alerts/fails
        if self.required_rotation:
            records['Approval'] = 'Pass'
            records['Code'] = [[]]*records.shape[0]

            # Alert from too much rotation at cycle 3, rd 2
            code = 1
            rd2_cyc3_rotation_high = self.criteria["Values"]["rd2_cyc3_rotation_high"]
            indicies = (records['Second Round']['Cycle 3+'] >= rd2_cyc3_rotation_high)
            records.loc[indicies, 'Approval'] = 'Alert'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

            # Alert from too many cycles
            code = 2
            cycles_high = self.criteria["Values"]["cycles_high"]
            indicies = (records['First Round']['# Cycles'] > cycles_high) | (
                records['Second Round']['# Cycles'] > cycles_high)
            records.loc[indicies, 'Approval'] = 'Alert'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

            # Alert from too much rotation
            code = 3
            perbolt_rotation_high = self.criteria["Values"]["perbolt_rotation_high"]
            indicies = (records["Total Rotation"] >= perbolt_rotation_high*self.required_rotation)
            records.loc[indicies, 'Approval'] = 'Alert'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

            # Alert from cycles not matching total
            code = 4
            indicies = ((rd1_total - records["First Round"]["Round Total"]).abs() >= buffer) | (
                (rd2_total - records["Second Round"]["Round Total"]).abs() >= buffer)
            records.loc[indicies, 'Approval'] = 'Alert'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

            # Alert from missing records in one round
            code = 5
            # diff_list = list(set(records1["BoltNo"]).symmetric_difference(records2["BoltNo"]))
            indicies = records["BoltNo"].isin(
                set(records1["BoltNo"]).symmetric_difference(records2["BoltNo"]))
            records.loc[indicies, 'Approval'] = 'Alert'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

            # Fail from too little rotation
            code = -1
            indicies = (records['Total Rotation'] < self.required_rotation)
            records.loc[indicies, 'Approval'] = 'Fail'
            records.loc[indicies, 'Code'] = records.loc[indicies,
                                                        'Code'].apply(lambda x: x + [code])

        else:
            # If we do not have a required rotation level, alert all bolts.
            records['Approval'] = "Alert"

        records = records.sort_values(by=["BoltNo"])
        self.records = records
        return records

    def __get_stats(self):
        """
        Determines the number of bolts rotated per cycle and the mean /
        standard devation of rotation for each round.

        Returns
        -------
        stats : dict
            Dict of two dataframes with the count and average/deviance bolt stats.

        """

        # Get all bolt rotation cycle columns
        cols = [col for col in self.records.columns if "Cycle" in col[1]
                and "s" not in col[1]]
        # Count number in each cycle that are experienced rotation
        count = self.records[cols].fillna(0).astype(bool).sum(axis=0)
        # Arrange into dataframe
        count = count.to_numpy()
        count = np.vstack([count[0:3], count[3:6]]).transpose()
        count = pd.DataFrame(data=count,
                             index=["Cycle 1", "Cycle 2", "Cycle 3+"],
                             columns=["First Round", "Second Round"])

        # Get the mean and standard deviation on the total rotations
        # Get the total rotation columns (1st total, 2nd total, 3rd total)
        totals = self.records[[
            col for col in self.records.columns if "Total" in col[1] or "Total" in col[0]]].to_numpy()
        # Calculate mean & SD
        totals = np.around(
            np.vstack([np.nanmean(totals, 0), np.nanstd(totals, 0)]), 1)
        # Arrange into dataframe
        totals = pd.DataFrame(data=totals,
                              index=["Mean Rotation", "Standard Deviation"],
                              columns=["First Round", "Second Round", "Total Rotation"])
        stats = {}
        stats["count"] = count
        stats["total"] = totals

        # Raise alerts if Mean is too high
        total_mean_high = self.criteria["Values"]["total_mean_high"]
        total_mean_veryhigh = self.criteria["Values"]["total_mean_veryhigh"]
        if self.required_rotation:
            if totals["Total Rotation"]["Mean Rotation"] < self.required_rotation:
                self.errors += "\nMean rotation is less than required rotation."
            elif totals["Total Rotation"]["Mean Rotation"] >= total_mean_veryhigh*self.required_rotation:
                self.errors += "\nMean rotation is excessively high."
            elif totals["Total Rotation"]["Mean Rotation"] >= total_mean_high*self.required_rotation:
                self.errors += "\nMean rotation is high."

        # Raise alters in SD is too high
        total_SD_high = self.criteria["Values"]["total_SD_high"]
        total_SD_veryhigh = self.criteria["Values"]["total_SD_veryhigh"]
        if totals["Total Rotation"]["Standard Deviation"] >= total_SD_veryhigh*self.required_rotation:
            self.errors += "\nStandard deviation of rotation is excessively high."
        elif totals["Total Rotation"]["Standard Deviation"] >= total_SD_high*self.required_rotation:
            self.errors += "\nStandard deviation of rotation is high."

        self.stats = stats
        return stats

    def __generate_response(self):
        """
        Uses the failed bolts and the canned responses to generate a response to 
        recommend repair actions to the site.

        Returns
        -------
        None.

        """
        self.response = ""
        pass

    def run(self):
        """
        Combined above function to run full analysis on flange

        Returns
        -------
        int
            output code.

        """

        # Declare position
        print(
            f"__Report: {self.location['project']}-{self.location['tower']}-{self.location['flange']} - ", end="")

        try:

            self.errors = ''

            # Data parsing
            # Get 1st and second round data
            self.xml_data["first"] = self.__get_data("first")
            self.xml_data["second"] = self.__get_data("second")

            # Check that data has been found
            if any([self.xml_data["first"]["headers"].empty, self.xml_data["second"]["headers"].empty,
                    self.xml_data["first"]["records"].empty, self.xml_data["second"]["records"].empty]):
                self.errors += (
                    "Required header or record data not determined.")

            # Evaluate headers
            self.headers = self.__eval_headers()

            # Evaluate bolts
            self.records = self.__eval_bolts()

            # Calculate stats
            self.stats = self.__get_stats()
            print(" Flange analysis complete.")

            # Write errors to a text file in flange path
            # if self.errors:
            #     with open(self.path+"\\Error.txt", "w") as writer:
            #         writer.write(self.errors)
            #     print("____Errors or alerts detected. Check error log file.")

            self.has_run = 1
            return 1
        except Exception as error:
            print(error)
            with open(self.path+"\\Error.txt", "w") as writer:
                writer.write(str(error))
            return 0
        pass
