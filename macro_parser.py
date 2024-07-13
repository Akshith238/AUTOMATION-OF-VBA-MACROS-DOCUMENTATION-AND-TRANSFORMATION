import logging
import os
import re
import win32com.client
import pythoncom

logger = logging.getLogger(__name__)

class MacroParser:
    def __init__(self):
        self.macro_code = ""
        self.global_variables = set()
        self.data_flow = {}

    def load_from_excel(self, file_path):
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            logger.info(f"Attempting to open file: {file_path}")
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            workbook = excel.Workbooks.Open(file_path)
            for component in workbook.VBProject.VBComponents:
                if component.Type in [1, 2, 3]:  # Modules, Class Modules, and Forms
                    code_module = component.CodeModule
                    lines = code_module.CountOfLines
                    if lines > 0:
                        self.macro_code += code_module.Lines(1, lines) + "\n\n"
            workbook.Close(SaveChanges=False)
            logger.info("Successfully loaded macro code from Excel file")
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            raise
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

    def parse_macros(self):
        self.analyze_global_variables()
        procedures = re.split(r'(Sub |Function )', self.macro_code)[1:]
        parsed_macros = []
        
        for i in range(0, len(procedures), 2):
            proc_type = procedures[i].strip()
            proc_code = procedures[i] + procedures[i+1]
            name = proc_code.split("(")[0].strip().split()[-1]
            parsed_macros.append(self.analyze_procedure(proc_type, name, proc_code))
        
        self.analyze_data_flow(parsed_macros)
        return parsed_macros

    def analyze_global_variables(self):
        self.global_variables = set(re.findall(r'Public\s+(\w+)', self.macro_code))

    def analyze_procedure(self, proc_type, name, code):
        args_match = re.search(r'\((.*?)\)', code)
        args = args_match.group(1) if args_match else ""
        
        return_type = ""
        if proc_type == "Function":
            return_type_match = re.search(r'As (\w+)', code)
            return_type = return_type_match.group(1) if return_type_match else "Variant"
        
        local_variables = set(re.findall(r'Dim\s+(\w+)', code))
        all_variables = local_variables.union(self.global_variables)
        
        variable_assignments = self.analyze_variable_assignments(code, all_variables)
        variable_usage = self.analyze_variable_usage(code, all_variables)
        
        return {
            'type': proc_type,
            'name': name,
            'arguments': args,
            'return_type': return_type,
            'local_variables': local_variables,
            'variable_assignments': variable_assignments,
            'variable_usage': variable_usage,
            'code': code
        }

    def analyze_variable_assignments(self, code, variables):
        assignments = {}
        for var in variables:
            assignments[var] = re.findall(rf'\b{var}\s*=\s*([^=\n]+)', code)
        return assignments

    def analyze_variable_usage(self, code, variables):
        usage = {}
        for var in variables:
            usage[var] = len(re.findall(r'\b' + var + r'\b', code))
        return usage

    def analyze_data_flow(self, parsed_macros):
        self.data_flow = {macro['name']: {'inputs': set(), 'outputs': set()} for macro in parsed_macros}
        
        for macro in parsed_macros:
            # Identify inputs (used variables that weren't assigned in this macro)
            inputs = set(macro['variable_usage'].keys()) - set(macro['variable_assignments'].keys())
            self.data_flow[macro['name']]['inputs'] = inputs

            # Identify outputs (assigned variables)
            outputs = set(macro['variable_assignments'].keys())
            self.data_flow[macro['name']]['outputs'] = outputs

            # Check for global variable modifications
            for var in self.global_variables:
                if var in outputs:
                    self.data_flow[macro['name']]['outputs'].add(f"global:{var}")

    def generate_markdown_documentation(self, parsed_macros):
        doc = []
        doc.append("# VBA Macro Analysis\n")

        doc.append("## Global Variables")
        for var in self.global_variables:
            doc.append(f"- `{var}`")
        doc.append("\n")

        for macro in parsed_macros:
            doc.append(f"## {macro['type']} {macro['name']}")
            doc.append(f"**Arguments:** {macro['arguments']}")
            if macro['type'] == 'Function':
                doc.append(f"**Return Type:** {macro['return_type']}")
            
            doc.append("### Local Variables")
            for var in macro['local_variables']:
                doc.append(f"- `{var}`")
            
            doc.append("### Variable Assignments")
            for var, assignments in macro['variable_assignments'].items():
                doc.append(f"- `{var}`: {', '.join(assignments)}")
            
            doc.append("### Variable Usage")
            for var, count in macro['variable_usage'].items():
                doc.append(f"- `{var}`: used {count} times")
            
            doc.append("### Data Flow")
            doc.append(f"**Inputs:** {', '.join(self.data_flow[macro['name']]['inputs'])}")
            doc.append(f"**Outputs:** {', '.join(self.data_flow[macro['name']]['outputs'])}")
            
            doc.append("### Code")
            doc.append("```vba")
            doc.append(macro['code'])
            doc.append("```")
            doc.append("\n")

        doc.append("## Overall Data Flow")
        for macro_name, flow in self.data_flow.items():
            doc.append(f"### {macro_name}")
            doc.append(f"**Inputs:** {', '.join(flow['inputs'])}")
            doc.append(f"**Outputs:** {', '.join(flow['outputs'])}")
        
        return "\n".join(doc)
    
    def infer_purpose(self, macro):
        name = macro['name'].lower()
        code = macro['code'].lower()
        if 'hello' in name:
            return "Displays a greeting message"
        elif 'add' in name and 'number' in name:
            return "Performs addition of two numbers"
        elif 'highlight' in name:
            return "Highlights cells based on a condition"
        elif 'create' in name and 'populate' in name:
            return "Creates and populates a new worksheet with data"
        elif 'calc' in name or 'calculate' in name:
            return "Performs calculations"
        elif 'update' in name:
            return "Updates data"
        elif 'get' in name or 'fetch' in name:
            return "Retrieves information"
        elif 'report' in name:
            return "Generates a report"
        elif 'validate' in name or 'check' in name:
            return "Validates data"
        else:
            return "Performs data processing"

    def explain_process(self, macro):
        code = macro['code'].lower()
        processes = []
        if 'msgbox' in code:
            processes.append("Displays a message box with a greeting")
        if '+' in code and macro['type'] == 'Function':
            processes.append("Adds two numbers together")
        if 'for each' in code:
            processes.append("Iterates through a range of cells")
        if 'for' in code:
            processes.append("Iterates through a series of items")
        if 'if' in code:
            processes.append("Makes decisions based on conditions")
        if 'color' in code or 'interior.color' in code:
            processes.append("Changes the color of cells")
        if 'worksheets.add' in code:
            processes.append("Creates a new worksheet")
        if 'cells' in code and '=' in code:
            processes.append("Populates cells with data")
        if 'font.bold' in code:
            processes.append("Formats cells as bold")
        if 'autofit' in code:
            processes.append("Adjusts column widths to fit content")
        if not processes:
            processes.append("Processes data")
        return ". ".join(processes)

    def explain_inputs(self, macro):
        inputs = [arg.strip() for arg in macro['arguments'].split(',')] if macro['arguments'] else []
        if 'range(' in macro['code'].lower():
            inputs.append("Specified range of cells")
        return f"Takes {', '.join(inputs) if inputs else 'no'} inputs"

    def explain_outputs(self, macro):
        outputs = []
        code = macro['code'].lower()
        if macro['type'] == 'Function':
            outputs.append(f"Returns a {macro['return_type']} value")
        if 'msgbox' in code:
            outputs.append("Displays a message to the user")
        if 'interior.color' in code:
            outputs.append("Modified cell colors")
        if 'worksheets.add' in code:
            outputs.append("New worksheet")
        if 'cells' in code and '=' in code:
            outputs.append("Populated cells with data")
        if not outputs:
            outputs.append("No direct outputs")
        return f"Produces {', '.join(outputs)}"

    def infer_business_impact(self, macro):
        name = macro['name'].lower()
        code = macro['code'].lower()
        if 'hello' in name:
            return "Provides a user-friendly interface element"
        elif 'add' in name and 'number' in name:
            return "Supports basic arithmetic operations in business calculations"
        elif 'highlight' in name:
            return "Enhances data visibility and aids in quick identification of specific information"
        elif 'create' in name and 'populate' in name:
            return "Automates data entry and report generation, improving efficiency and consistency"
        elif 'report' in name or 'summary' in name:
            return "Aids in decision-making by providing summarized information"
        elif 'calc' in name or 'calculate' in name:
            return "Ensures accurate financial or operational calculations"
        elif 'update' in name or 'modify' in name:
            return "Maintains data integrity and currency"
        elif 'validate' in name or 'check' in name:
            return "Ensures data quality and compliance"
        else:
            return "Supports business operations through data processing"

    def extract_functional_logic(self, parsed_macros):
        logic_explanations = []
        for macro in parsed_macros:
            logic_explanations.append(self.explain_macro_logic(macro))
        return logic_explanations

    def explain_macro_logic(self, macro):
        explanation = {
            'name': macro['name'],
            'type': macro['type'],
            'purpose': self.infer_purpose(macro),
            'inputs': self.explain_inputs(macro),
            'process': self.explain_process(macro),
            'outputs': self.explain_outputs(macro),
            'business_impact': self.infer_business_impact(macro)
        }
        return explanation

    # def generate_functional_documentation(self, logic_explanations):
    #     doc = []
    #     doc.append("# Functional Logic Explanation of VBA Macros\n")

    #     for explanation in logic_explanations:
    #         doc.append(f"## {explanation['type']} {explanation['name']}\n")
    #         doc.append(f"**Purpose:** {explanation['purpose']}\n")
    #         doc.append(f"**Inputs:** {explanation['inputs']}\n")
    #         doc.append(f"**Process:** {explanation['process']}\n")
    #         doc.append(f"**Outputs:** {explanation['outputs']}\n")
    #         doc.append(f"**Business Impact:** {explanation['business_impact']}\n")
    #         doc.append(f"**Enhanced Explanation:** {explanation['enhanced_explanation']}\n")
    #         doc.append("\n")

    #     return "\n".join(doc)
    
    def generate_functional_documentation(self, logic_explanations):
        doc = []
        doc.append("# Functional Logic Explanation of VBA Macros\n")

        for explanation in logic_explanations:
            doc.append(explanation)
            doc.append("\n---\n")  # Add a separator between macro explanations

        return "\n".join(doc)
    
# Usage
if __name__ == "__main__":
    parser = MacroParser()
    parser.load_from_excel("C:/Users/jenis/Downloads/Book1.xlsm")
    parsed_macros = parser.parse_macros()
    logic_explanations = parser.extract_functional_logic(parsed_macros)
    functional_documentation = parser.generate_functional_documentation(logic_explanations)
    print(functional_documentation)