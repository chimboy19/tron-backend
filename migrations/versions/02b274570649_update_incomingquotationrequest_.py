"""Update IncomingQuotationRequest:  customer_id

Revision ID: 02b274570649
Revises: 
Create Date: 2025-11-22 19:16:49.701466

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '02b274570649'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('incoming_quotation_requests', schema=None) as batch_op:
        batch_op.add_column(sa.Column('customer_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            'fk_incomingreq_customer_id',   # <-- ADD FK NAME HERE
            'customers',
            ['customer_id'],
            ['id']
        )


def downgrade():
    with op.batch_alter_table('incoming_quotation_requests', schema=None) as batch_op:
        batch_op.drop_constraint(
            'fk_incomingreq_customer_id',   # <-- MUST MATCH EXACT NAME
            type_='foreignkey'
        )
        batch_op.drop_column('customer_id')
