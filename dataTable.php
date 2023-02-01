<?php

namespace Yolva\Reports;

use Bitrix\Catalog\ProductTable;
use CCrmCompany;
use CIBlock;
use CIBlockElement;
use Error;
use Yolva\Helper\EntityFactory;

class DataTable
{
    private $entityId;
    private $entityTypeId;
    public function __construct($entityTypeId, $entityId)
    {
        $this->entityTypeId = $entityTypeId;
        $this->entityId = $entityId;
    }
    private function getEntityData($entityTypeId, $entityId)
    {
        $item = EntityFactory::getItem($entityTypeId, $entityId);
        $data = $item->getData();
        return $data;
    }
    /**
     * @deprecated this method retrieve products by UF_CRM_2_1660569878
     */
    private function getProductsByFieldValues($productsId): array
    {
        $entityTypeId = EntityFactory::getEntityTypeIdByCode('QUOTE_PRODUCT');
        foreach ($productsId as $id) {
            $item = EntityFactory::getItem($entityTypeId, $id);
            $dataCollection[] = $item->getData();
        }
        return $dataCollection;
    }
    private function getQuoteProductsData(): array
    {
        $entityTypeId = EntityFactory::getEntityTypeIdByCode('QUOTE_PRODUCT');
        $parentId = 'PARENT_ID_' . $this->entityTypeId;
        $items = EntityFactory::getItemsByParameters(
            $entityTypeId,
            [
                'filter' =>  [$parentId => $this->entityId]
            ]
        );
        foreach ($items as $item) {
            $data[] = $item->getData();
        }
        return $data;
    }
    private function getGoodsByQuoteProductId($quoteProductId)
    {
        $entityTypeId = EntityFactory::getEntityTypeIdByCode('QUOTE_PRODUCT');
        $goodsRaw = \CCrmProductRow::GetList(array(), array('OWNER_ID' => $quoteProductId, 'ONWER_TYPE' => dechex($entityTypeId)));
        while ($good = $goodsRaw->GetNext()) {
            foreach ($this->getPropertiesProduct($good['PRODUCT_ID']) as $key => $propertyValue) {
                $good[$key] = $propertyValue;
            }
            $additionalParameters = $this->getAdditionalParameters($good['PRODUCT_ID']);
            $good['PARAMETRS'] = $additionalParameters;
            $goods[] = $good;
        }
        return $goods;
    }
    private function getAdditionalParameters($productId)
    {
        return ProductTable::getById($productId)->Fetch();
    }
    private function getVariations($cml2_link)
    {
        $iblock_offer = CIBlock::GetList(array(), array('CODE' => "crm_offers"))->Fetch()['ID'];
        $list = CIBlockElement::GetList(array(), array('IBLOCK_ID' => $iblock_offer, 'PROPERTY_CML2_LINK' => $cml2_link));
        while ($item = $list->GetNextElement()) {

            $properties = $item->GetProperties(false, array('CODE' => 'TIP_TSENY_FKA81Z'));
            $fields = $item->GetFields();
            $variations = CatalogGetPriceTable($fields['ID']);
            $result[] = [
                'PRICE_TYPE' => $properties['TIP_TSENY_FKA81Z']['VALUE'],
                'PRICE' => $variations['MATRIX'][0]['PRICE'][1]
            ];
        }
        return $result;
    }
    private function getPropertiesProduct($productId)
    {
        $iblock_offer = CIBlock::GetList(array(), array('CODE' => "crm_offers"))->Fetch()['ID'];
        $offers = CIBlockElement::GetList(array(), array('IBLOCK_ID' => $iblock_offer, 'ID' => $productId));
        while ($offer = $offers->GetNextElement()) {
            $linkMainProductId = $offer->GetProperties(false, array("CODE" => "CML2_LINK"))['CML2_LINK']['VALUE'];
            $variations = $this->getVariations($linkMainProductId);
            yield 'VARIATIONS' => $variations;

            $mainProducts = CIBlockElement::GetList(array(), array('ID' => $linkMainProductId));
            while ($product = $mainProducts->GetNextElement()) {
                $properties = $product->GetProperties();
                foreach ($properties as $property) {
                    yield $property['CODE'] => $property['VALUE'];
                }
            }
        }
    }
    public function getQuote()
    {
        $quote = $this->getEntityData($this->entityTypeId, $this->entityId);
        $companyId = $quote['COMPANY_ID'];
        if (!isset($companyId)) throw new Error('Компания в смете не указана');
        $company = CCrmCompany::GetByID($companyId);
        $quote['COMPANY'] = $company;
        return $quote;
    }
    private function GetQuoteProductsGoods(): array
    {
        $products = $this->getQuoteProductsData();
        foreach ($products as $quoteProduct) {
            $quoteProduct['GOODS'] = $this->getGoodsByQuoteProductId($quoteProduct['ID']);
            $result[] = $quoteProduct;
        }
        return $result;
    }
    public function getData(): array
    {
        return $this->GetQuoteProductsGoods();
    }
}
