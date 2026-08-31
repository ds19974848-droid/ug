import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from src.models import Base, CostItem
from src.cost_item_service import find_cost_item_reference


@pytest.fixture()
def db_session():
    engine = create_engine('sqlite:///:memory:')
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    session = Session()
    try:
        yield session
    finally:
        session.close()


def test_find_cost_item_exact(db_session):
    item = CostItem(source_key='k1', item_name='测试项 A', unit='m2')
    db_session.add(item)
    db_session.commit()
    found, score = find_cost_item_reference(db_session, '测试项 A', unit='m2')
    assert found is not None
    assert found.item_name == '测试项 A'
    assert score == 1.0


def test_find_cost_item_unit_mismatch(db_session):
    item = CostItem(source_key='k2', item_name='测试项 B', unit='套')
    db_session.add(item)
    db_session.commit()
    found, score = find_cost_item_reference(db_session, '测试项 B', unit='m2')
    assert found is not None
    assert found.item_name == '测试项 B'
    assert 0.7 <= score <= 0.95


def test_find_cost_item_fuzzy(db_session):
    item = CostItem(source_key='k3', item_name='水泥 C30', unit='袋')
    db_session.add(item)
    db_session.commit()
    found, score = find_cost_item_reference(db_session, 'C 30 水泥', unit='袋')
    assert found is not None
    assert found.item_name == '水泥 C30'
    assert score > 0.7
