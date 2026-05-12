<template>
    <div id="functionBox">
        <!-- 标题 -->
        <h1 class="font" style="margin-bottom: 15px">自定义函数描述</h1>

        <!-- 文档 -->
        <n-scrollbar style="max-height: 83vh">
            <n-collapse style="padding: 0px 15px; word-wrap: break-word" accordion>
                <!-- 0.仪表相关处理函数 分割线 -->
                <n-divider>
                    <span class="font" style="color: yellowgreen">特殊相关处理函数</span>
                </n-divider>
                <n-collapse-item title="getZHLJWorkTime 根据六化建实名制打卡时间算出具体时间-漳州南化项目" name="0-1">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>timeStr</td>
                                <td>输入打卡时间</td>
                                <td>
                                    <p>06:38,</p>
                                    <p>11:14,</p>
                                    <p>13:05,</p>
                                    <p>18:41</p>
                                </td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getZHLJWorkTime('06:38,11:14,13:05,18:41')</p>
                        <div>
                            <p>输出:</p>
                            <p>天:01 (9小时为一天)</p>
                            <p>时:01 (剩余小时)</p>
                            <p>分:12 (剩余分钟)</p>
                            <p>次:04 (打卡次数)</p>
                        </div>
                    </n-alert>
                </n-collapse-item>

                <!-- 1.字符串相关处理函数 分割线 -->
                <n-divider>
                    <span class="font" style="color: yellowgreen">字符串相关处理函数</span>
                </n-divider>
                <n-collapse-item title="cardIDForAge 根据身份证号返回年龄" name="1-1">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>id</td>
                                <td>输入身份证号码</td>
                                <td>没人不知道</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.cardIDForAge('******19991012****')</p>
                    </n-alert>
                </n-collapse-item>

                <!-- 2.仪表相关处理函数 分割线 -->
                <n-divider>
                    <span class="font" style="color: yellowgreen">仪表相关处理函数</span>
                </n-divider>
                <n-collapse-item title="getInputThreeData 根据仪表量程范围返回三点数值" name="2-1">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>range</td>
                                <td>输入仪表量程</td>
                                <td>0~1.6Mpa</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>fixed</td>
                                <td>保留小数点</td>
                                <td>2</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getInputThreeData('0~1.6Mpa', 2)</p>
                        <p>输出: 0.00 / 0.80 / 1.60</p>
                    </n-alert>
                </n-collapse-item>
                <n-collapse-item title="getInputFiveData 根据仪表量程范围返回五点数值" name="2-2">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>range</td>
                                <td>输入仪表量程</td>
                                <td>0~1.6Mpa</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>fixed</td>
                                <td>保留小数点</td>
                                <td>2</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getInputFiveData('0~1.6Mpa', 2)</p>
                        <p>输出: 0.00 / 0.40 / 0.80 / 1.20 / 1.60</p>
                    </n-alert>
                </n-collapse-item>
                <n-collapse-item title="getInputUnit 根据仪表量程范围返回对应单位" name="2-3">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>range</td>
                                <td>输入仪表量程</td>
                                <td>0~1.6Mpa</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getInputUnit('0~1.6Mpa')</p>
                        <p>输出: Mpa</p>
                    </n-alert>
                </n-collapse-item>
                <n-collapse-item title="getValveStandardThreeData 根据FC/FO类型返回阀门标准三点行程" name="2-4">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>range</td>
                                <td>输入阀门行程</td>
                                <td>0~25mm</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>type</td>
                                <td>输入阀门类型</td>
                                <td>FO</td>
                            </tr>
                            <tr>
                                <td>3</td>
                                <td>fixed</td>
                                <td>保留小数点</td>
                                <td>2</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getValveStandardThreeData('0~25mm', 'FO', 2)</p>
                        <p>输出: 25.00 / 12.50 / 0.00</p>
                    </n-alert>
                </n-collapse-item>
                <n-collapse-item title="getValveAllowableError 根据精确度及行程返回阀门允许误差" name="2-5">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>accuracy</td>
                                <td>输入精确度</td>
                                <td>1.0级/0.1%</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>itinerary</td>
                                <td>输入阀门行程</td>
                                <td>0~25mm</td>
                            </tr>
                            <tr>
                                <td>3</td>
                                <td>fixed</td>
                                <td>保留小数点</td>
                                <td>3</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>{{ NameSpace.toString() }}.getValveAllowableError('1.0级', '0~25mm', 2)</p>
                        <p>输出: 0.250mm</p>
                    </n-alert>
                </n-collapse-item>
                <n-collapse-item title="getValveError 根据阀门实测行程返回误差" name="2-6">
                    <n-table :single-line="false" size="small" class="fontCenter">
                        <thead>
                            <tr>
                                <th>序号</th>
                                <th>参数</th>
                                <th>描述</th>
                                <th>举例</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>1</td>
                                <td>range</td>
                                <td>输入行程范围</td>
                                <td>0~25mm</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>data</td>
                                <td>输入实测行程</td>
                                <td>Excel Area</td>
                            </tr>
                            <tr>
                                <td>3</td>
                                <td>direction</td>
                                <td>输入作用方向</td>
                                <td>FO</td>
                            </tr>
                            <tr>
                                <td>4</td>
                                <td>fixed</td>
                                <td>保留小数点</td>
                                <td>2</td>
                            </tr>
                        </tbody>
                    </n-table>
                    <n-alert title="实例" type="info" :show-icon="false" style="margin-top: 15px">
                        <p>data=[25.00,12.50,0.00,25.00,12.50,0.00]</p>
                        <p>{{ NameSpace.toString() }}.getValveError('0~25mm',data,'FO',2)</p>
                        <p>输出: mm(误差单位) 0.00(正1) 0.00(正2) 0.00(正3) 0.00(反1) 0.00(反2) 0.00(反3) mm(回差单位) 0.00(回差1) 0.00(回差2) 0.00(回差3)</p>
                    </n-alert>
                </n-collapse-item>
            </n-collapse>
        </n-scrollbar>
    </div>
</template>

<script setup>
import { NameSpace } from '../utils/utils';
</script>

<style scoped>
#functionBox {
    width: 100%;
    height: 100%;
    margin-bottom: 0px;
}

.font {
    display: block;
    text-align: center;
    font-family: '得意黑', sans-serif;
}

.fontCenter {
    text-align: center;
}
</style>
