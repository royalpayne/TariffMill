"""
Database Migration Tool for DerivativeMill
Migrates data between SQLite and PostgreSQL databases
"""

import sys
import argparse
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from database import DatabaseManager, migrate_sqlite_to_postgresql, POSTGRES_AVAILABLE


def test_postgresql_connection(host: str, port: int, database: str, user: str, password: str) -> bool:
    """Test if we can connect to the PostgreSQL server"""
    if not POSTGRES_AVAILABLE:
        print("ERROR: psycopg2 is not installed. Run: pip install psycopg2-binary")
        return False
    
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=host,
            port=port,
            database=database,
            user=user,
            password=password
        )
        conn.close()
        print(f"✓ Successfully connected to PostgreSQL at {host}:{port}/{database}")
        return True
    except Exception as e:
        print(f"✗ Failed to connect to PostgreSQL: {e}")
        return False


def setup_postgresql_database(host: str, port: int, database: str, user: str, password: str):
    """Create the database schema on PostgreSQL"""
    db_manager = DatabaseManager()
    db_manager.configure_postgresql(host, port, database, user, password)
    db_manager.init_database()
    print("✓ PostgreSQL schema initialized")
    return db_manager


def migrate_to_postgresql(sqlite_path: Path, host: str, port: int, database: str, user: str, password: str):
    """Full migration from SQLite to PostgreSQL"""
    print("\n" + "=" * 60)
    print("DerivativeMill: SQLite to PostgreSQL Migration")
    print("=" * 60)
    
    # Verify SQLite file exists
    if not sqlite_path.exists():
        print(f"ERROR: SQLite database not found: {sqlite_path}")
        return False
    
    print(f"\nSource: {sqlite_path}")
    print(f"Target: postgresql://{user}@{host}:{port}/{database}")
    
    # Test PostgreSQL connection
    print("\n1. Testing PostgreSQL connection...")
    if not test_postgresql_connection(host, port, database, user, password):
        return False
    
    # Setup PostgreSQL schema
    print("\n2. Initializing PostgreSQL schema...")
    pg_manager = setup_postgresql_database(host, port, database, user, password)
    
    # Migrate data
    print("\n3. Migrating data...")
    migrate_sqlite_to_postgresql(sqlite_path, pg_manager)
    
    print("\n" + "=" * 60)
    print("Migration completed successfully!")
    print("=" * 60)
    print("\nNext steps:")
    print("1. Update db_config.json to use 'postgresql' as the type")
    print("2. Restart DerivativeMill to use the new database")
    print("3. Verify data integrity in the PostgreSQL database")
    
    return True


def switch_to_sqlite(db_path: str = None):
    """Switch back to SQLite database"""
    db_manager = DatabaseManager()
    db_manager.configure_sqlite(db_path)
    db_manager.init_database()
    print(f"✓ Switched to SQLite database: {db_path or 'default'}")


def show_current_config():
    """Display current database configuration"""
    db_manager = DatabaseManager()
    config = db_manager.get_config()
    
    print("\nCurrent Database Configuration:")
    print("-" * 40)
    print(f"Type: {config.get('type', 'unknown')}")
    
    if config.get('type') == 'sqlite':
        sqlite_cfg = config.get('sqlite', {})
        print(f"Path: {sqlite_cfg.get('path', 'N/A')}")
    elif config.get('type') == 'postgresql':
        pg_cfg = config.get('postgresql', {})
        print(f"Host: {pg_cfg.get('host', 'N/A')}")
        print(f"Port: {pg_cfg.get('port', 'N/A')}")
        print(f"Database: {pg_cfg.get('database', 'N/A')}")
        print(f"User: {pg_cfg.get('user', 'N/A')}")


def main():
    parser = argparse.ArgumentParser(
        description="DerivativeMill Database Migration Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Show current configuration
  python migrate_db.py --status
  
  # Migrate SQLite to PostgreSQL
  python migrate_db.py --migrate \\
      --sqlite-path ./Resources/derivativemill.db \\
      --pg-host localhost --pg-port 5432 \\
      --pg-database derivativemill --pg-user postgres --pg-password secret
  
  # Switch back to SQLite
  python migrate_db.py --use-sqlite
  
  # Test PostgreSQL connection
  python migrate_db.py --test-pg \\
      --pg-host localhost --pg-port 5432 \\
      --pg-database derivativemill --pg-user postgres --pg-password secret
"""
    )
    
    parser.add_argument('--status', action='store_true', help='Show current database configuration')
    parser.add_argument('--migrate', action='store_true', help='Migrate SQLite to PostgreSQL')
    parser.add_argument('--use-sqlite', action='store_true', help='Switch to SQLite database')
    parser.add_argument('--test-pg', action='store_true', help='Test PostgreSQL connection')
    
    # SQLite options
    parser.add_argument('--sqlite-path', type=str, help='Path to SQLite database file')
    
    # PostgreSQL options
    parser.add_argument('--pg-host', type=str, default='localhost', help='PostgreSQL host')
    parser.add_argument('--pg-port', type=int, default=5432, help='PostgreSQL port')
    parser.add_argument('--pg-database', type=str, default='derivativemill', help='PostgreSQL database name')
    parser.add_argument('--pg-user', type=str, default='postgres', help='PostgreSQL username')
    parser.add_argument('--pg-password', type=str, default='', help='PostgreSQL password')
    
    args = parser.parse_args()
    
    if args.status:
        show_current_config()
    
    elif args.test_pg:
        test_postgresql_connection(
            args.pg_host, args.pg_port, args.pg_database, 
            args.pg_user, args.pg_password
        )
    
    elif args.migrate:
        sqlite_path = Path(args.sqlite_path) if args.sqlite_path else Path(__file__).parent / "Resources" / "derivativemill.db"
        migrate_to_postgresql(
            sqlite_path,
            args.pg_host, args.pg_port, args.pg_database,
            args.pg_user, args.pg_password
        )
    
    elif args.use_sqlite:
        switch_to_sqlite(args.sqlite_path)
    
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
