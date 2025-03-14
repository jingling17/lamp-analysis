import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pathlib import Path

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

class LampAnalysis:
    def __init__(self, file_path):
        self.df = pd.read_excel(file_path)
        print("Excel文件的列名：", self.df.columns)  # 添加这行来查看列名
        self.price_ranges = [0, 100, 200, 300, 400, 500, 800, 1000, float('inf')]
        self.price_labels = ['0-100', '100-200', '200-300', '300-400', 
                           '400-500', '500-800', '800-1000', '1000+']
        
    def add_price_range(self):
        """添加价格区间列"""
        self.df['价格区间'] = pd.cut(self.df['价格'], bins=self.price_ranges, labels=self.price_labels)
        
    def analyze_total_sales(self):
        """分析全年销售额和销量"""
        total_sales = self.df['销售额'].sum()
        total_volume = self.df['销量'].sum()
        
        print("\n=== 全年销售数据分析 ===")
        print(f"总销售额：{total_sales:,.2f} 元")
        print(f"总销量：{total_volume:,.0f} 件")
        
        # 创建柱状图
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        ax1.bar(['销售额'], [total_sales])
        ax1.set_title('总销售额')
        ax1.set_ylabel('金额（元）')
        
        ax2.bar(['销量'], [total_volume])
        ax2.set_title('总销量')
        ax2.set_ylabel('数量（件）')
        
        plt.tight_layout()
        plt.savefig('total_sales_analysis.png')
        plt.close()
        
    def analyze_price_range_distribution(self):
        """分析价位段分布"""
        price_range_stats = self.df.groupby('价格区间').agg({
            '销售额': 'sum',
            '销量': 'sum'
        })
        
        print("\n=== 价位段分布分析 ===")
        for price_range in self.price_labels:
            if price_range in price_range_stats.index:
                sales = price_range_stats.loc[price_range, '销售额']
                volume = price_range_stats.loc[price_range, '销量']
                print(f"\n价位段 {price_range}:")
                print(f"销售额：{sales:,.2f} 元 ({sales/price_range_stats['销售额'].sum()*100:.1f}%)")
                print(f"销量：{volume:,.0f} 件 ({volume/price_range_stats['销量'].sum()*100:.1f}%)")
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        # 销售额占比饼图
        ax1.pie(price_range_stats['销售额'], labels=price_range_stats.index,
                autopct='%1.1f%%')
        ax1.set_title('各价位段销售额占比')
        
        # 销量占比饼图
        ax2.pie(price_range_stats['销量'], labels=price_range_stats.index,
                autopct='%1.1f%%')
        ax2.set_title('各价位段销量占比')
        
        plt.tight_layout()
        plt.savefig('price_range_distribution.png')
        plt.close()
        
    def analyze_top_brands_by_price_range(self):
        """分析每个价位段TOP5品牌"""
        result = {}
        for price_range in self.price_labels:
            range_data = self.df[self.df['价格区间'] == price_range]
            top_brands = range_data.groupby('品牌').agg({
                '销售额': 'sum',
                '销量': 'sum'
            }).sort_values('销售额', ascending=False).head(5)
            
            result[price_range] = top_brands
            
        return result
    
    def analyze_top_products_by_price_range(self):
        """分析每个价位段TOP5商品"""
        result = {}
        for price_range in self.price_labels:
            range_data = self.df[self.df['价格区间'] == price_range]
            top_products = range_data.sort_values('销售额', ascending=False).head(5)[
                ['商品标题', '商品链接', '销售额', '销量']]
            result[price_range] = top_products
            
        return result
    
    def analyze_brand_market_share(self):
        """分析品牌市场占比"""
        brand_stats = self.df.groupby('品牌').agg({
            '销售额': 'sum',
            '销量': 'sum'
        }).sort_values('销售额', ascending=False).head(10)
        
        print("\n=== TOP10品牌市场占比分析 ===")
        total_sales = self.df['销售额'].sum()
        total_volume = self.df['销量'].sum()
        
        for brand in brand_stats.index:
            sales = brand_stats.loc[brand, '销售额']
            volume = brand_stats.loc[brand, '销量']
            print(f"\n品牌：{brand}")
            print(f"销售额：{sales:,.2f} 元 ({sales/total_sales*100:.1f}%)")
            print(f"销量：{volume:,.0f} 件 ({volume/total_volume*100:.1f}%)")
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
        
        # 销售额占比
        ax1.pie(brand_stats['销售额'], labels=brand_stats.index, autopct='%1.1f%%')
        ax1.set_title('TOP10品牌销售额占比')
        
        # 销量占比
        ax2.pie(brand_stats['销量'], labels=brand_stats.index, autopct='%1.1f%%')
        ax2.set_title('TOP10品牌销量占比')
        
        plt.tight_layout()
        plt.savefig('brand_market_share.png')
        plt.close()
        
    def analyze_top_brands_price_distribution(self):
        """分析TOP5品牌在各价位段的分布"""
        top_5_brands = self.df.groupby('品牌')['销售额'].sum().nlargest(5).index
        
        print("\n=== TOP5品牌价位段分布分析 ===")
        for brand in top_5_brands:
            brand_data = self.df[self.df['品牌'] == brand].groupby('价格区间').agg({
                '销售额': 'sum',
                '销量': 'sum'
            })
            
            total_brand_sales = brand_data['销售额'].sum()
            total_brand_volume = brand_data['销量'].sum()
            
            print(f"\n品牌：{brand}")
            print(f"总销售额：{total_brand_sales:,.2f} 元")
            print(f"总销量：{total_brand_volume:,.0f} 件")
            print("\n各价位段分布：")
            
            for price_range in self.price_labels:
                if price_range in brand_data.index:
                    sales = brand_data.loc[price_range, '销售额']
                    volume = brand_data.loc[price_range, '销量']
                    print(f"\n  {price_range}:")
                    print(f"  销售额：{sales:,.2f} 元 ({sales/total_brand_sales*100:.1f}%)")
                    print(f"  销量：{volume:,.0f} 件 ({volume/total_brand_volume*100:.1f}%)")
            
        brand_price_stats = pd.DataFrame()
        for brand in top_5_brands:
            brand_data = self.df[self.df['品牌'] == brand].groupby('价格区间').agg({
                '销售额': 'sum',
                '销量': 'sum'
            })
            brand_price_stats[brand] = brand_data['销售额']
            
        # 创建堆叠柱状图
        ax = brand_price_stats.plot(kind='bar', stacked=True, figsize=(12, 6))
        plt.title('TOP5品牌各价位段销售额分布')
        plt.xlabel('价格区间')
        plt.ylabel('销售额')
        plt.legend(title='品牌')
        plt.xticks(rotation=45)
        
        plt.tight_layout()
        plt.savefig('top_brands_price_distribution.png')
        plt.close()

    def save_analysis_to_excel(self):
        """将分析结果保存到Excel文件"""
        # 创建一个空的DataFrame列表，用于存储所有数据
        all_data = []
        
        # 总销售数据
        total_sales = self.df['销售额'].sum()
        total_volume = self.df['销量'].sum()
        total_data = pd.DataFrame({
            '分析类型': ['总体数据'] * 2,
            '指标': ['总销售额', '总销量'],
            '数值': [total_sales, total_volume]
        })
        all_data.append(total_data)
        
        # 价位段分布数据
        price_range_stats = self.df.groupby('价格区间').agg({
            '销售额': 'sum',
            '销量': 'sum'
        }).reset_index()
        price_range_stats['分析类型'] = '价位段分布'
        price_range_stats['销售额占比'] = price_range_stats['销售额'] / total_sales * 100
        price_range_stats['销量占比'] = price_range_stats['销量'] / total_volume * 100
        all_data.append(price_range_stats)
        
        # 各价位段TOP5品牌数据
        for price_range in self.price_labels:
            range_data = self.df[self.df['价格区间'] == price_range]
            range_sales = range_data['销售额'].sum()
            range_volume = range_data['销量'].sum()
            
            top_brands = range_data.groupby('品牌').agg({
                '销售额': 'sum',
                '销量': 'sum'
            }).reset_index().sort_values('销售额', ascending=False).head(5)
            
            top_brands['分析类型'] = f'{price_range}价位TOP5品牌'
            top_brands['价格区间'] = price_range
            top_brands['销售额占总体比例'] = top_brands['销售额'] / total_sales * 100
            top_brands['销量占总体比例'] = top_brands['销量'] / total_volume * 100
            top_brands['销售额占价位段比例'] = top_brands['销售额'] / range_sales * 100
            top_brands['销量占价位段比例'] = top_brands['销量'] / range_volume * 100
            
            all_data.append(top_brands)
        
        # 各价位段TOP5商品数据
        for price_range in self.price_labels:
            range_data = self.df[self.df['价格区间'] == price_range]
            range_sales = range_data['销售额'].sum()
            range_volume = range_data['销量'].sum()
            
            top_products = range_data.sort_values('销售额', ascending=False).head(5)[
                ['商品标题', '商品链接', '销售额', '销量']]
            top_products['分析类型'] = f'{price_range}价位TOP5商品'
            top_products['价格区间'] = price_range
            top_products['销售额占总体比例'] = top_products['销售额'] / total_sales * 100
            top_products['销量占总体比例'] = top_products['销量'] / total_volume * 100
            top_products['销售额占价位段比例'] = top_products['销售额'] / range_sales * 100
            top_products['销量占价位段比例'] = top_products['销量'] / range_volume * 100
            
            all_data.append(top_products)
        
        # TOP10品牌市场占比
        brand_stats = self.df.groupby('品牌').agg({
            '销售额': 'sum',
            '销量': 'sum'
        }).reset_index().sort_values('销售额', ascending=False).head(10)
        
        brand_stats['分析类型'] = 'TOP10品牌市场占比'
        brand_stats['销售额占比'] = brand_stats['销售额'] / total_sales * 100
        brand_stats['销量占比'] = brand_stats['销量'] / total_volume * 100
        all_data.append(brand_stats)
        
        # TOP5品牌价位段分布
        top_5_brands = self.df.groupby('品牌')['销售额'].sum().nlargest(5).index
        for brand in top_5_brands:
            brand_data = self.df[self.df['品牌'] == brand].groupby('价格区间').agg({
                '销售额': 'sum',
                '销量': 'sum'
            }).reset_index()
            
            brand_total_sales = brand_data['销售额'].sum()
            brand_total_volume = brand_data['销量'].sum()
            
            brand_data['分析类型'] = f'{brand}品牌价位段分布'
            brand_data['品牌'] = brand
            brand_data['销售额占品牌总额比例'] = brand_data['销售额'] / brand_total_sales * 100
            brand_data['销量占品牌总量比例'] = brand_data['销量'] / brand_total_volume * 100
            brand_data['销售额占总体比例'] = brand_data['销售额'] / total_sales * 100
            brand_data['销量占总体比例'] = brand_data['销量'] / total_volume * 100
            
            all_data.append(brand_data)
        
        # 合并所有数据并保存到Excel
        final_data = pd.concat(all_data, ignore_index=True)
        with pd.ExcelWriter('台灯销售分析报告.xlsx') as writer:
            final_data.to_excel(writer, sheet_name='销售分析报告', index=False)

# 在main函数中添加调用
def main():
    # 获取target目录下的Excel文件
    target_dir = Path('target')
    excel_files = list(target_dir.glob('*.xlsx'))
    
    if not excel_files:
        print("未找到Excel文件！")
        return
        
    analyzer = LampAnalysis(excel_files[0])
    analyzer.add_price_range()
    
    # 执行各项分析
    analyzer.analyze_total_sales()
    analyzer.analyze_price_range_distribution()
    
    # 输出TOP5品牌分析结果
    top_brands = analyzer.analyze_top_brands_by_price_range()
    for price_range, brands in top_brands.items():
        print(f"\n{price_range}价位段TOP5品牌：")
        print(brands)
    
    # 输出TOP5商品分析结果
    top_products = analyzer.analyze_top_products_by_price_range()
    for price_range, products in top_products.items():
        print(f"\n{price_range}价位段TOP5商品：")
        print(products)
    
    analyzer.analyze_brand_market_share()
    analyzer.analyze_top_brands_price_distribution()
    analyzer.save_analysis_to_excel()  # 添加这一行

if __name__ == "__main__":
    main()