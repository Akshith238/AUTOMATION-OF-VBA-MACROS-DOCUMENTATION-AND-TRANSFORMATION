import re

class MacroQualityAnalyzer:
    def __init__(self, macro_parser):
        self.macro_parser = macro_parser

    def analyze_macros(self):
        parsed_macros = self.macro_parser.parse_macros()
        analysis_results = []
        efficient_macros = []
        inefficient_macros = []

        for macro in parsed_macros:
            analysis = self.analyze_macro(macro)
            if analysis['issues']:
                analysis['efficiency'] = 'Inefficient'
                inefficient_macros.append(analysis)
            else:
                analysis['efficiency'] = 'Efficient'
                efficient_macros.append(analysis)
            analysis_results.append(analysis)

        return analysis_results, efficient_macros, inefficient_macros

    def analyze_macro(self, macro):
        analysis = {
            'name': macro['name'],
            'type': macro['type'],
            'issues': [],
            'suggestions': []
        }

        self.check_variable_naming(macro, analysis)
        self.check_code_complexity(macro, analysis)
        self.check_error_handling(macro, analysis)
        self.check_performance_issues(macro, analysis)
        self.check_code_duplication(macro, analysis)
        self.check_unused_variables(macro, analysis)
        self.check_option_explicit(macro, analysis)

        return analysis

    def check_variable_naming(self, macro, analysis):
        poor_names = re.findall(r'\b([a-z]{1,2})\b', macro['code'])
        if poor_names:
            analysis['issues'].append(f"Poor variable naming: {', '.join(set(poor_names))}")
            analysis['suggestions'].append("Use descriptive variable names")

    def check_code_complexity(self, macro, analysis):
        nested_levels = max(len(line.strip()) - len(line.strip().lstrip()) for line in macro['code'].split('\n'))
        if nested_levels > 4:
            analysis['issues'].append(f"High code complexity: {nested_levels} levels of nesting")
            analysis['suggestions'].append("Reduce nesting by extracting code into separate procedures")

    def check_error_handling(self, macro, analysis):
        if 'On Error' not in macro['code']:
            analysis['issues'].append("No error handling")
            analysis['suggestions'].append("Implement error handling using 'On Error GoTo' statements")

    def check_performance_issues(self, macro, analysis):
        if 'Select' in macro['code'] or 'Activate' in macro['code']:
            analysis['issues'].append("Use of Select/Activate which can be slow")
            analysis['suggestions'].append("Avoid using Select/Activate, instead use direct references")

        if 'For ' in macro['code'] and '.Value' in macro['code']:
            analysis['issues'].append("Possible slow cell value access in loop")
            analysis['suggestions'].append("Consider using array for bulk operations instead of accessing individual cells")

    def check_code_duplication(self, macro, analysis):
        lines = macro['code'].split('\n')
        for i in range(len(lines)):
            for j in range(i+1, len(lines)):
                if lines[i] == lines[j] and lines[i].strip() and not lines[i].strip().startswith("'"):
                    analysis['issues'].append(f"Duplicate code: '{lines[i].strip()}'")
                    analysis['suggestions'].append("Extract duplicate code into a separate procedure")
                    break

    def check_unused_variables(self, macro, analysis):
        declared_vars = set(re.findall(r'Dim\s+(\w+)', macro['code']))
        used_vars = set(re.findall(r'\b(\w+)\b', macro['code']))
        unused_vars = declared_vars - used_vars
        if unused_vars:
            analysis['issues'].append(f"Unused variables: {', '.join(unused_vars)}")
            analysis['suggestions'].append("Remove unused variables to improve code clarity")

    def check_option_explicit(self, macro, analysis):
        if 'Option Explicit' not in self.macro_parser.macro_code:
            analysis['issues'].append("'Option Explicit' not used")
            analysis['suggestions'].append("Add 'Option Explicit' at the top of the module to enforce variable declaration")

    def generate_analysis_report(self, analysis_results, efficient_macros, inefficient_macros):
        report = ["# VBA Macro Quality and Efficiency Analysis\n"]

        report.append("## Efficient Macros\n")
        for macro in efficient_macros:
            report.append(f"- {macro['type']} {macro['name']}\n")
        
        report.append("\n## Inefficient Macros\n")
        for macro in inefficient_macros:
            report.append(f"- {macro['type']} {macro['name']} \n")

        report.append("\n## Detailed Analysis\n")
        for analysis in analysis_results:
            report.append(f"### {analysis['type']} {analysis['name']}\n")
            report.append(f"Efficiency: {analysis['efficiency']}\n")
            
            if analysis['issues']:
                report.append("#### Issues:\n")
                for issue in analysis['issues']:
                    report.append(f"- {issue}\n")
            else:
                report.append("#### Issues:\nNo issues found.\n")

            if analysis['suggestions']:
                report.append("\n#### Suggestions:\n")
                for suggestion in analysis['suggestions']:
                    report.append(f"- {suggestion}\n")
            else:
                report.append("\n#### Suggestions:\nNo suggestions available.\n")

            report.append("\n---\n")

        return "\n".join(report)
    
from macro_parser import MacroParser

if __name__ == "__main__":
    # Initialize parser and load Excel file
    parser = MacroParser()
    parser.load_from_excel("C:/Users/jenis/Downloads/Book1.xlsm")
    
    # Initialize analyzer and perform macro analysis
    analyzer = MacroQualityAnalyzer(parser)
    analysis_results, efficient_macros, inefficient_macros = analyzer.analyze_macros()
    
    # Generate and print analysis report
    analysis_report = analyzer.generate_analysis_report(analysis_results, efficient_macros, inefficient_macros)
    print(analysis_report)
