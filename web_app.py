import streamlit as st
import pandas as pd
from lamp_analysis import LampAnalysis
import os

def main():
    st.set_page_config(page_title="台灯销售数据分析工具", layout="wide")
    
    st.title("台灯销售数据分析工具")
    
    # 文件上传
    uploaded_file = st.file_uploader("选择Excel文件", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # 保存上传的文件
            with open("temp.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            if st.button("开始分析"):
                with st.spinner("正在分析数据..."):
                    # 运行分析
                    analyzer = LampAnalysis("temp.xlsx")
                    analyzer.add_price_range()
                    analyzer.analyze_total_sales()
                    analyzer.analyze_price_range_distribution()
                    analyzer.analyze_brand_market_share()
                    analyzer.analyze_top_brands_price_distribution()
                    analyzer.save_analysis_to_excel()
                    
                    # 提供下载链接
                    with open("台灯销售分析报告.xlsx", "rb") as file:
                        st.download_button(
                            label="下载分析报告",
                            data=file,
                            file_name="台灯销售分析报告.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # 显示生成的图片
                    if os.path.exists("total_sales_analysis.png"):
                        st.image("total_sales_analysis.png", caption="总销售分析")
                    if os.path.exists("price_range_distribution.png"):
                        st.image("price_range_distribution.png", caption="价位段分布")
                    if os.path.exists("brand_market_share.png"):
                        st.image("brand_market_share.png", caption="品牌市场占比")
                    if os.path.exists("top_brands_price_distribution.png"):
                        st.image("top_brands_price_distribution.png", caption="TOP5品牌价位段分布")
                    
                    st.success("分析完成！")
                    
        except Exception as e:
            st.error(f"分析过程中出现错误：{str(e)}")
        finally:
            # 清理临时文件
            if os.path.exists("temp.xlsx"):
                os.remove("temp.xlsx")

if __name__ == "__main__":
    main()