export default function ExcelAutomationPlatform() {
  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-6xl mx-auto space-y-6">
        <div className="bg-white rounded-3xl shadow-sm p-8 border">
          <h1 className="text-3xl font-bold mb-2">Excel 自动化网页平台（原型方案）</h1>
          <p className="text-gray-600 leading-7">
            你现在的 Excel 自动化逻辑（XLOOKUP、MAXIFS、随访判断、购药分析、药房匹配等）完全可以迁移成网页系统。
            后续只需要上传固定模板 Excel，系统即可自动清洗、计算、输出结果。
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white rounded-2xl shadow-sm p-6 border">
            <div className="text-xl font-semibold mb-3">① 上传模板</div>
            <p className="text-gray-600 text-sm leading-6 mb-4">
              上传固定格式 Excel 文件（销售底表 / 随访底表）。
            </p>
            <div className="border-2 border-dashed rounded-2xl p-10 text-center text-gray-500">
              点击上传 Excel
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm p-6 border">
            <div className="text-xl font-semibold mb-3">② 自动计算</div>
            <ul className="space-y-2 text-sm text-gray-600 leading-6">
              <li>• 自动匹配患者</li>
              <li>• 自动识别超期患者</li>
              <li>• 自动判断二次随访</li>
              <li>• 自动计算购药间隔</li>
              <li>• 自动生成提醒文案</li>
            </ul>
          </div>

          <div className="bg-white rounded-2xl shadow-sm p-6 border">
            <div className="text-xl font-semibold mb-3">③ 输出结果</div>
            <ul className="space-y-2 text-sm text-gray-600 leading-6">
              <li>• 在线查看结果</li>
              <li>• 导出 Excel</li>
              <li>• 导出随访任务</li>
              <li>• 导出 BI 数据</li>
              <li>• 对接飞书/企微</li>
            </ul>
          </div>
        </div>

        <div className="bg-white rounded-3xl shadow-sm p-8 border">
          <h2 className="text-2xl font-bold mb-6">你当前这个 Excel 可以迁移的核心能力</h2>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {[
              'XLOOKUP 患者匹配',
              'MAXIFS 末次随访判断',
              '购药时间间隔分析',
              '患者超期识别',
              '本月未购药原因提醒',
              '药房简称自动转换',
              '随访状态判断',
              '自动生成提醒话术'
            ].map((item) => (
              <div key={item} className="rounded-2xl border p-5 bg-gray-50">
                <div className="font-medium">{item}</div>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-3xl shadow-sm p-8 border">
          <h2 className="text-2xl font-bold mb-6">推荐系统结构（适合你现在的业务）</h2>

          <div className="space-y-5">
            <div className="rounded-2xl border p-5">
              <div className="font-semibold text-lg mb-2">前端网页</div>
              <p className="text-gray-600 leading-7 text-sm">
                上传 Excel、查看计算结果、筛选患者、导出数据。
                可以做成和你现在 BI 大屏一样的科技感风格。
              </p>
            </div>

            <div className="rounded-2xl border p-5">
              <div className="font-semibold text-lg mb-2">自动化计算引擎</div>
              <p className="text-gray-600 leading-7 text-sm">
                将你现在 Excel 的公式逻辑改成程序逻辑。
                后续不再依赖复杂 Excel 公式。
              </p>
            </div>

            <div className="rounded-2xl border p-5">
              <div className="font-semibold text-lg mb-2">模板识别系统</div>
              <p className="text-gray-600 leading-7 text-sm">
                以后只要上传固定格式模板，系统自动识别字段并完成计算。
              </p>
            </div>

            <div className="rounded-2xl border p-5">
              <div className="font-semibold text-lg mb-2">自动提醒模块（后期）</div>
              <p className="text-gray-600 leading-7 text-sm">
                可直接对接飞书 / 企业微信，自动推送超期患者与随访任务。
              </p>
            </div>
          </div>
        </div>

        <div className="bg-blue-50 rounded-3xl p-8 border border-blue-100">
          <h2 className="text-2xl font-bold mb-4">最适合你的方案（非常关键）</h2>

          <div className="space-y-4 text-gray-700 leading-8">
            <p>
              你现在不建议直接做“完全自由的 Excel 平台”。
            </p>

            <p>
              最适合你的方式是：
              <span className="font-semibold">固定模板 + 自动计算引擎</span>
            </p>

            <p>
              也就是：
              以后所有门店都上传统一模板，系统自动跑你现在的规则。
            </p>

            <p>
              这样稳定性最高，维护成本最低，而且最适合医药零售业务。
            </p>
          </div>
        </div>

        <div className="bg-white rounded-3xl shadow-sm p-8 border">
          <h2 className="text-2xl font-bold mb-6">后续还能继续升级</h2>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 text-sm text-gray-700">
            <div className="rounded-2xl border p-5">✓ 自动生成患者随访任务</div>
            <div className="rounded-2xl border p-5">✓ 自动统计门店执行率</div>
            <div className="rounded-2xl border p-5">✓ 自动生成 PPT 数据</div>
            <div className="rounded-2xl border p-5">✓ 自动识别异常患者</div>
            <div className="rounded-2xl border p-5">✓ 对接企业微信提醒</div>
            <div className="rounded-2xl border p-5">✓ BI 大屏实时展示</div>
          </div>
        </div>

        <div className="bg-black text-white rounded-3xl p-8">
          <h2 className="text-2xl font-bold mb-4">你这个表其实已经接近“小型业务系统”了</h2>

          <p className="leading-8 text-gray-300">
            你现在很多逻辑已经不是普通 Excel，而是：
            患者数据管理 + 自动随访判断 + 超期监控 + 数据清洗 + 自动提醒。
            所以完全值得做成网页系统。
          </p>
        </div>
      </div>
    </div>
  )
}