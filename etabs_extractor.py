import pandas as pd
import os
import sys

def get_beam_results(edb_path):
    """
    Connects to ETABS, opens the model, extracts beam forces and required reinforcement,
    and returns them as pandas DataFrames.
    """
    try:
        import comtypes.client
    except ImportError:
        raise Exception("comtypes library not installed. Please run on Windows.")

    try:
        # Create ETABS object
        myETABSObject = comtypes.client.CreateObject("CSI.ETABS.API.ETABSObject")
        
        # Start ETABS
        myETABSObject.ApplicationStart()
        
        # Get SapModel object
        SapModel = myETABSObject.SapModel
        
        # Initialize model
        SapModel.InitializeNewModel()
        
        # Open the specific file
        ret = SapModel.File.OpenFile(edb_path)
        if ret != 0:
            raise Exception(f"Could not open file: {edb_path}")
            
        # Run Analysis
        SapModel.Analyze.RunAnalysis()
        
        # Deselect all cases and combos for output to save time, then select a few common ones
        SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        # You can customize this to select specific combinations, e.g. "Env"
        # SapModel.Results.Setup.SetComboSelectedForOutput("Env")
        
        # --- 1. Extract Beam Forces ---
        # 2 = Frame objects
        NumberResults = 0
        Obj = []
        ObjSta = []
        Elm = []
        ElmSta = []
        LoadCase = []
        StepType = []
        StepNum = []
        P = []
        V2 = []
        V3 = []
        T = []
        M2 = []
        M3 = []
        
        ret = SapModel.Results.FrameForce("All", 0, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)
        
        if ret[0] == 0:
            # Successfully extracted forces
            forces_data = {
                "Frame": ret[3],
                "Station": ret[4],
                "LoadCase": ret[7],
                "P": ret[10],
                "V2": ret[11],
                "V3": ret[12],
                "T": ret[13],
                "M2": ret[14],
                "M3": ret[15]
            }
            df_forces = pd.DataFrame(forces_data)
        else:
            df_forces = pd.DataFrame()

        # --- 2. Extract Required Reinforcement (Concrete Design) ---
        # First, run the design
        SapModel.DesignConcrete.StartDesign()
        
        # Get summary results for beams
        # Since we don't know exact names, we can loop through all frame objects, 
        # check if it's a beam, and get design results.
        NumberItems = 0
        FrameName = []
        MyOption = 2 # 2 = Beam
        NumberItems, FrameName, ret = SapModel.FrameObj.GetNameList(NumberItems, FrameName)
        
        rebar_data = []
        
        for name in FrameName:
            # Check frame design procedure
            prog_type, ret2 = SapModel.FrameObj.GetDesignProcedure(name)
            if prog_type == 2: # 2 = Concrete Frame Design
                # We can try to get beam results
                try:
                    num_results, obj, sta, top_rebar, bot_rebar, v_rebar, t_rebar, ret3 = SapModel.DesignConcrete.GetSummaryResultsBeam(name, 0)
                    if ret3 == 0:
                        for i in range(num_results):
                            rebar_data.append({
                                "Frame": obj[i],
                                "Station": sta[i],
                                "TopRebar": top_rebar[i],
                                "BotRebar": bot_rebar[i],
                                "ShearRebar": v_rebar[i],
                                "TorsionRebar": t_rebar[i]
                            })
                except Exception:
                    pass
        
        df_rebar = pd.DataFrame(rebar_data)

        # Close ETABS
        myETABSObject.ApplicationExit(False)
        
        return df_forces, df_rebar

    except Exception as e:
        # Ensure we try to close ETABS if an error occurs
        try:
            myETABSObject.ApplicationExit(False)
        except:
            pass
        raise e

if __name__ == "__main__":
    # Test block
    print("ETABS Extractor initialized.")
