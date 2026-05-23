from docxtpl import DocxTemplate
import pandas as pd
import os

def generate_report(template_path, output_path, project_name, engineer_name, df_forces, df_rebar):
    """
    Takes the Word template, replaces tags with project info and ETABS data, 
    and saves the generated document.
    """
    # 1. Initialize Template
    doc = DocxTemplate(template_path)

    # 2. Format DataFrames for docxtpl
    # docxtpl expects lists of dictionaries for tables
    
    # Format forces table: round values to 2 decimal places
    forces_list = []
    if not df_forces.empty:
        # We might not want to dump thousands of rows, so let's just dump top 50 
        # or aggregate them (e.g., max forces per frame). For now, we take top 50.
        # In a real scenario, you'd filter for Max/Min combos.
        df_forces_sample = df_forces.head(50).copy()
        
        # Round numeric columns
        numeric_cols = df_forces_sample.select_dtypes(include=['float64']).columns
        df_forces_sample[numeric_cols] = df_forces_sample[numeric_cols].round(2)
        
        forces_list = df_forces_sample.to_dict('records')

    # Format rebar table
    rebar_list = []
    if not df_rebar.empty:
        df_rebar_sample = df_rebar.head(50).copy()
        numeric_cols = df_rebar_sample.select_dtypes(include=['float64']).columns
        df_rebar_sample[numeric_cols] = df_rebar_sample[numeric_cols].round(2)
        
        rebar_list = df_rebar_sample.to_dict('records')

    # 3. Create context dictionary
    context = {
        'project_name': project_name,
        'engineer_name': engineer_name,
        'forces_list': forces_list,
        'rebar_list': rebar_list
    }

    # 4. Render and Save
    doc.render(context)
    doc.save(output_path)
    
    return output_path

if __name__ == "__main__":
    # Test block
    print("Report Generator initialized.")
