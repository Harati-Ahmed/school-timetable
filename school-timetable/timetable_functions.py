from typing import List, Dict, Any, Optional
from datetime import datetime

class TimetableError(Exception):
    def __init__(self, message: str, error_type: str = "SYSTEM_ERROR", details: Dict = None):
        super().__init__(message)
        self.type = error_type
        self.details = details or {}
        self.timestamp = datetime.now()

def validate_config_data(config_data: Dict[str, List[List[str]]]) -> List[str]:
    """
    Validates configuration data for teachers, classes, and subjects.
    
    Args:
        config_data: Dictionary containing teachers, classes, and subjects data
        
    Returns:
        List of error messages. Empty list if validation passes.
    """
    errors = []
    teacher_ids = set()
    class_ids = set()
    subject_ids = set()
    
    # Validate teachers
    for index, teacher in enumerate(config_data.get('teachers', [])):
        row_num = index + 2  # Add 2 to account for 1-based index and header row
        
        if len(teacher) < 3:
            errors.append(f"Invalid teacher data at row {row_num}: missing required fields")
            continue
            
        teacher_id, name, subject = teacher[:3]
        
        # Check for duplicate ID
        if teacher_id in teacher_ids:
            errors.append(f"Duplicate Teacher ID '{teacher_id}' found at row {row_num}")
        else:
            teacher_ids.add(teacher_id)
        
        # Validate ID format
        if not teacher_id or not teacher_id.startswith('T') or not teacher_id[1:].isdigit() or len(teacher_id) != 4:
            errors.append(f"Invalid Teacher ID format at row {row_num}. Must be 'T' followed by 3 digits")
        
        # Validate name and subject
        if not name or not name.strip():
            errors.append(f"Missing Teacher Name at row {row_num}")
        if not subject or not subject.strip():
            errors.append(f"Missing Subject at row {row_num}")
    
    # Validate classes
    for index, class_data in enumerate(config_data.get('classes', [])):
        row_num = index + 2
        
        if len(class_data) < 2:
            errors.append(f"Invalid class data at row {row_num}: missing required fields")
            continue
            
        class_id, name = class_data[:2]
        
        # Check for duplicate ID
        if class_id in class_ids:
            errors.append(f"Duplicate Class ID '{class_id}' found at row {row_num}")
        else:
            class_ids.add(class_id)
        
        # Validate ID format
        if not class_id or not class_id.startswith('C') or not class_id[1:].isdigit() or len(class_id) != 4:
            errors.append(f"Invalid Class ID format at row {row_num}. Must be 'C' followed by 3 digits")
        
        # Validate name
        if not name or not name.strip():
            errors.append(f"Missing Class Name at row {row_num}")
    
    # Validate subjects
    for index, subject in enumerate(config_data.get('subjects', [])):
        row_num = index + 2
        
        if len(subject) < 2:
            errors.append(f"Invalid subject data at row {row_num}: missing required fields")
            continue
            
        subject_id, name = subject[:2]
        
        # Check for duplicate ID
        if subject_id in subject_ids:
            errors.append(f"Duplicate Subject ID '{subject_id}' found at row {row_num}")
        else:
            subject_ids.add(subject_id)
        
        # Validate ID format
        if not subject_id or not subject_id.startswith('S') or not subject_id[1:].isdigit() or len(subject_id) != 4:
            errors.append(f"Invalid Subject ID format at row {row_num}. Must be 'S' followed by 3 digits")
        
        # Validate name
        if not name or not name.strip():
            errors.append(f"Missing Subject Name at row {row_num}")
    
    # Validate subject references
    valid_subjects = {subject[1] for subject in config_data.get('subjects', [])}
    for index, teacher in enumerate(config_data.get('teachers', [])):
        if len(teacher) >= 3:
            subject = teacher[2]
            if subject not in valid_subjects:
                errors.append(f"Invalid subject '{subject}' for teacher at row {index + 2}")
    
    return errors

def calculate_teacher_workload(teacher_data: List[List[str]], break_col: int, lunch_col: int) -> Dict[str, int]:
    """
    Calculates the workload (number of periods) for each teacher.
    
    Args:
        teacher_data: List of teacher rows with their schedule
        break_col: Column index for break period (1-based, e.g. 7 for column G)
        lunch_col: Column index for lunch period (1-based, e.g. 11 for column K)
        
    Returns:
        Dictionary mapping teacher names to their total periods
    """
    workload = {}
    
    for row in teacher_data[1:]:  # Skip header row
        if len(row) < 2:  # Skip invalid rows
            continue
            
        teacher_name = row[1]
        if not teacher_name:  # Skip empty rows
            continue
        
        # Count periods (excluding break and lunch)
        total_periods = 0
        periods = row[3:]  # Get all periods starting from first period
        
        print(f"\nProcessing {teacher_name}:")
        print(f"Periods: {periods}")
        
        # Count non-empty periods excluding break and lunch
        for i, cell in enumerate(periods):
            actual_col = i + 4  # Convert to actual column number (1-based)
            # Skip empty cells, break period, and lunch period
            if cell and cell.strip():
                if actual_col == break_col:
                    print(f"Column {actual_col}: {cell} (skipped - break)")
                elif actual_col == lunch_col:
                    print(f"Column {actual_col}: {cell} (skipped - lunch)")
                else:
                    print(f"Column {actual_col}: {cell} (counted)")
                    total_periods += 1
            else:
                print(f"Column {actual_col}: {cell} (skipped - empty)")
        
        workload[teacher_name] = total_periods
    
    return workload

def validate_timetable_conflicts(teacher_data: List[List[str]], class_data: List[List[str]]) -> List[str]:
    """
    Validates the timetable for various conflicts.
    
    Args:
        teacher_data: List of teacher rows with their schedule
        class_data: List of class rows with their schedule
        
    Returns:
        List of conflict messages. Empty list if no conflicts found.
    """
    conflicts = []
    
    # Create period-wise mappings
    teacher_class_count = {}  # {teacher: {period: [classes]}}
    class_teacher_count = {}  # {class: {period: [teachers]}}
    
    print("\nProcessing teacher data:")
    # Process teacher data
    for row in teacher_data[1:]:  # Skip header
        if len(row) < 4:  # Need at least teacher name and one period
            continue
        
        teacher_name = row[1]
        if not teacher_name:
            continue
        
        print(f"\nTeacher: {teacher_name}")
        if teacher_name not in teacher_class_count:
            teacher_class_count[teacher_name] = {}
            
        # Get periods starting from first period (index 3)
        periods = row[3:]
        for i, class_entry in enumerate(periods):
            period = i + 1  # Convert to 1-based period number
            
            if not class_entry or class_entry == 'Break' or class_entry == 'Lunch':
                print(f"Period {period}: {class_entry} (skipped)")
                continue
            
            print(f"Period {period}: {class_entry}")
            
            # Split multiple classes in the same cell
            classes = [c.strip() for c in class_entry.split(',')]
            
            if period not in teacher_class_count[teacher_name]:
                teacher_class_count[teacher_name][period] = []
            
            teacher_class_count[teacher_name][period].extend(classes)
            
            # Check for multiple classes in the same period
            if len(set(teacher_class_count[teacher_name][period])) > 1:
                unique_classes = sorted(set(teacher_class_count[teacher_name][period]))
                conflict_msg = (
                    f"Teacher '{teacher_name}' is assigned multiple classes in period {period}: "
                    f"{', '.join(unique_classes)}"
                )
                print(f"Conflict detected: {conflict_msg}")
                conflicts.append(conflict_msg)
    
    print("\nProcessing class data:")
    # Process class data
    for row in class_data[1:]:  # Skip header
        if len(row) < 3:  # Need at least class name and one period
            continue
        
        class_name = row[1]
        if not class_name:
            continue
        
        print(f"\nClass: {class_name}")
        if class_name not in class_teacher_count:
            class_teacher_count[class_name] = {}
            
        # Get periods starting from first period (index 2 for class data)
        periods = row[2:]
        for i, teacher_entry in enumerate(periods):
            period = i + 1  # Convert to 1-based period number
            
            if not teacher_entry or teacher_entry == 'Break' or teacher_entry == 'Lunch':
                print(f"Period {period}: {teacher_entry} (skipped)")
                continue
            
            print(f"Period {period}: {teacher_entry}")
            
            # Split multiple teachers in the same cell
            teachers = [t.strip() for t in teacher_entry.split(',')]
            
            if period not in class_teacher_count[class_name]:
                class_teacher_count[class_name][period] = []
            
            class_teacher_count[class_name][period].extend(teachers)
            
            # Check for multiple teachers in the same period
            if len(set(class_teacher_count[class_name][period])) > 1:
                unique_teachers = sorted(set(class_teacher_count[class_name][period]))
                conflict_msg = (
                    f"Class '{class_name}' has multiple teachers in period {period}: "
                    f"{', '.join(unique_teachers)}"
                )
                print(f"Conflict detected: {conflict_msg}")
                conflicts.append(conflict_msg)
    
    return conflicts 