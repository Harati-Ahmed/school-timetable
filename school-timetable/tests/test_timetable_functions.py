import pytest
from timetable_functions import (
    validate_config_data,
    calculate_teacher_workload,
    validate_timetable_conflicts,
    TimetableError
)

# Test data fixtures
@pytest.fixture
def valid_config_data():
    return {
        'teachers': [
            ['T001', 'John Doe', 'Mathematics'],
            ['T002', 'Jane Smith', 'English'],
            ['T003', 'Bob Wilson', 'Science']
        ],
        'classes': [
            ['C001', 'Grade 1A'],
            ['C002', 'Grade 1B'],
            ['C003', 'Grade 2A']
        ],
        'subjects': [
            ['S001', 'Mathematics'],
            ['S002', 'English'],
            ['S003', 'Science']
        ]
    }

@pytest.fixture
def invalid_config_data():
    return {
        'teachers': [
            ['T001', 'John Doe', 'Mathematics'],
            ['T001', 'Jane Smith', 'English'],  # Duplicate ID
            ['TINV', '', 'Invalid']  # Invalid format and empty name
        ],
        'classes': [
            ['C001', 'Grade 1A'],
            ['CINV', ''],  # Invalid format and empty name
            ['C001', 'Grade 1B']  # Duplicate ID
        ],
        'subjects': [
            ['S001', 'Mathematics'],
            ['SINV', ''],  # Invalid format and empty name
            ['S001', 'English']  # Duplicate ID
        ]
    }

@pytest.fixture
def teacher_schedule_data():
    return [
        ['SI', 'Teacher Name', 'Subject', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        ['1', 'John Doe', 'Mathematics', 'Grade 1A', 'Grade 1B', '', 'Break', 'Grade 2A', '', '', 'Lunch', 'Grade 1A', '', ''],
        ['2', 'Jane Smith', 'English', '', 'Grade 2A', 'Grade 1A', 'Break', '', 'Grade 1B', '', 'Lunch', '', 'Grade 2A', '']
    ]

@pytest.fixture
def class_schedule_data():
    return [
        ['SI', 'Class', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        ['1', 'Grade 1A', 'John Doe / Mathematics', '', 'Jane Smith / English', 'Break', '', '', '', 'Lunch', 'John Doe / Mathematics', '', ''],
        ['2', 'Grade 1B', '', 'John Doe / Mathematics', '', 'Break', '', 'Jane Smith / English', '', 'Lunch', '', '', '']
    ]

def test_validate_config_data_valid(valid_config_data):
    errors = validate_config_data(valid_config_data)
    assert len(errors) == 0

def test_validate_config_data_invalid(invalid_config_data):
    errors = validate_config_data(invalid_config_data)
    assert len(errors) > 0
    # Check for specific error messages
    assert any('Duplicate Teacher ID' in error for error in errors)
    assert any('Invalid Teacher ID format' in error for error in errors)
    assert any('Missing Teacher Name' in error for error in errors)
    assert any('Duplicate Class ID' in error for error in errors)
    assert any('Invalid Class ID format' in error for error in errors)
    assert any('Missing Class Name' in error for error in errors)
    assert any('Duplicate Subject ID' in error for error in errors)
    assert any('Invalid Subject ID format' in error for error in errors)
    assert any('Missing Subject Name' in error for error in errors)

def test_calculate_teacher_workload(teacher_schedule_data):
    print("\nTeacher schedule data:")
    for row in teacher_schedule_data:
        print(row)
    
    workload = calculate_teacher_workload(teacher_schedule_data, break_col=7, lunch_col=11)
    print("\nCalculated workload:", workload)
    
    # John Doe's schedule: ['1', 'John Doe', 'Mathematics', 'Grade 1A', 'Grade 1B', '', 'Break', 'Grade 2A', '', '', 'Lunch', 'Grade 1A', '', '']
    # Should have 4 periods: Grade 1A (period 1), Grade 1B (period 2), Grade 2A (period 4), Grade 1A (period 7)
    assert workload['John Doe'] == 4  # 4 periods assigned
    assert workload['Jane Smith'] == 4  # 4 periods assigned

def test_validate_timetable_conflicts_no_conflicts(teacher_schedule_data, class_schedule_data):
    conflicts = validate_timetable_conflicts(teacher_schedule_data, class_schedule_data)
    assert len(conflicts) == 0

def test_validate_timetable_conflicts_with_conflicts():
    # Create test data with conflicts
    conflicting_teacher_data = [
        ['SI', 'Teacher Name', 'Subject', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        ['1', 'John Doe', 'Mathematics', 'Grade 1A', 'Grade 1B', 'Grade 2A', 'Break', 'Grade 2A', 'Grade 1A,Grade 2A', 'Grade 1B', 'Lunch', 'Grade 1A', '', ''],  # Multiple classes in period 5
        ['2', 'Jane Smith', 'English', 'Grade 2A', 'Grade 2A,Grade 1B', 'Grade 1A', 'Break', '', 'Grade 1B', 'Grade 2A', 'Lunch', '', 'Grade 2A', '']  # Multiple classes in period 2
    ]
    
    conflicting_class_data = [
        ['SI', 'Class', '1', '2', '3', 'Break', '4', '5', '6', 'Lunch', '7', '8', '9'],
        ['1', 'Grade 1A', 'John Doe / Mathematics', 'Jane Smith / English,John Doe / Mathematics', 'Jane Smith / English', 'Break', '', 'John Doe / Mathematics', 'John Doe / Mathematics', 'Lunch', 'John Doe / Mathematics', '', ''],  # Multiple teachers in period 2
        ['2', 'Grade 1B', 'Jane Smith / English', 'John Doe / Mathematics', '', 'Break', '', 'Jane Smith / English', 'John Doe / Mathematics', 'Lunch', '', '', '']  # Multiple teachers in period 6
    ]
    
    conflicts = validate_timetable_conflicts(conflicting_teacher_data, conflicting_class_data)
    print("\nDetected conflicts:", conflicts)  # Debug print
    assert len(conflicts) > 0, "No conflicts detected"
    assert any('multiple classes' in conflict for conflict in conflicts), "No multiple classes conflicts detected"
    assert any('multiple teachers' in conflict for conflict in conflicts), "No multiple teachers conflicts detected"

def test_timetable_error():
    error = TimetableError("Test error", "TEST_ERROR", {"detail": "test"})
    assert error.type == "TEST_ERROR"
    assert error.details == {"detail": "test"}
    assert str(error) == "Test error"
    assert hasattr(error, 'timestamp') 