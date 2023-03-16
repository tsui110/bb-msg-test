import aiomysql
import asyncio

loops = asyncio.get_event_loop()


# 定义一个执行语句的
def dosql(sql):
    task = loops.create_task(connect_mysql(sql))
    loops.run_until_complete(task)
    return task.result()


# 定义一个数据库实例
async def connect_mysql(sql=None):
    conn = await aiomysql.connect(host='127.0.0.1', port=3306,
                                  user='bigbear', password='bigbearpc', db='bigbearpc',
                                  cursorclass=aiomysql.cursors.DictCursor,
                                  loop=loops)
    if sql is None:
        pass
    else:
        cur = await conn.cursor()
        await cur.execute(sql)
        # print(cur.description)
        r = await cur.fetchall()
        await cur.close()
        conn.close()
        return r
