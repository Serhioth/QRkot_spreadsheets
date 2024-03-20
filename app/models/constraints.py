from sqlalchemy import CheckConstraint


check_description_length = CheckConstraint(
    'LENGTH(description) >= 1', name='check_description_length'
)
check_full_amount_is_positive = CheckConstraint(
    'full_amount > 0', name='check_full_amount_is_positive'
)
