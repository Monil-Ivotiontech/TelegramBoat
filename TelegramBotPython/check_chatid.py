import asyncio
import asyncpg
import os
from dotenv import load_dotenv

load_dotenv()

async def check_chatid(chat_id: int):
    conn = await asyncpg.connect(
        host=os.getenv('DB_HOST'),
        port=int(os.getenv('DB_PORT', 5432)),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_NAME'),
    )
    
    print(f"\n🔍 Checking authorization for Chat ID: {chat_id}\n")
    print("=" * 60)
    
    # Check MT5 managers
    print("\n📱 MT5 Managers (for PHONE/EMAIL bot access):")
    mt5_query = """
        SELECT c."Login", c.chatid, ba.bottype, cd.commandaccess
        FROM mt5_manager_chatid c
        LEFT JOIN mt5_manager_chatid_detail cd ON c."Login" = cd."Login"
        LEFT JOIN mt5_manager_bot_access ba ON c."Login" = ba."Login"
        WHERE c.chatid = $1 AND c.utilflag = 'active'
    """
    mt5_rows = await conn.fetch(mt5_query, chat_id)
    if mt5_rows:
        for row in mt5_rows:
            print(f"   ✅ Login: {row['Login']}, Bot: {row['bottype']}, Access: {row['commandaccess']}")
    else:
        print("   ❌ No MT5 manager access found for this Chat ID")
    
    # Check Venus managers
    print("\n📧 Venus Managers:")
    venus_query = """
        SELECT vm.id, vm.chatid, vmba.bottype, vm.commandaccess
        FROM venus_manager vm
        LEFT JOIN venus_manager_bot_access vmba ON vm.id = vmba.managerid
        WHERE vm.chatid = $1 AND vm.utilflag = 'active'
    """
    venus_rows = await conn.fetch(venus_query, chat_id)
    if venus_rows:
        for row in venus_rows:
            print(f"   ✅ Manager ID: {row['id']}, Bot: {row['bottype']}, Access: {row['commandaccess']}")
    else:
        print("   ❌ No Venus manager access found for this Chat ID")
    
    print("\n" + "=" * 60)
    
    await conn.close()

# Replace with your actual chat ID
YOUR_CHAT_ID = 7760526578  

asyncio.run(check_chatid(YOUR_CHAT_ID))
