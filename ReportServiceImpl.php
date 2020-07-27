<?php

namespace Modules\OMM\Service;

use App\Service\BaseService;
use DB;
use Auth;
use Maatwebsite\Excel\Facades\Excel;
use App\Service\Core\ConfigurationServiceImpl as ConfigurationService;
use App\Service\Core\UserServiceImpl as UserService;

/**
 * @author Eunjung Kim
 * @file ReportServiceImpl.php
 * @package Modules\OMM\Service
 * @brief ReportService.
 */
class ReportServiceImpl extends BaseService implements ReportService
{
    /**
     * Upload Excel for RMA Order Items
     *
     * @author Eunjung Kim - 05/28/2020 - updated for header validation - 06/26/2020
     * @param $dataArray
     * @param $uploadFile
     * @return Json Array
     */
    public function uploadExcelForRmaOrderItemBulkInsert(array $dataArray, $uploadFile)
    {
        ini_set('memory_limit', '512M');
        ini_set('max_execution_time', '1900');

        config([
            'excel.import.force_sheets_collection' => 'true',
            'excel.import.heading' => 'numeric',
            'excel.import.startRow' => 2
        ]);

        $connection = DB::connection('mysql');
        try {
            //Read the first two sheets, skip top 3 rows, and add to an array
            $uploadFileDataList = Excel::selectSheetsByIndex(0, 1)->load($uploadFile->path())
                ->formatDates(true)
                ->skipRows(2)
                ->toarray();

                $currUser       = Auth::user();
                $userService    = new UserService();
                $userCode       = $userService->findUserByEmail($currUser->email)->user_code;
                $systemId       = $dataArray['system_id'];
    
                $configurationService   = new ConfigurationService();
                $newStatusRmaOrder      = $configurationService->findCommonCodeItemByCode($systemId, 'rma_order_status', 'new')->code_id ?? 1;

            //If Data exists,
            if (count($uploadFileDataList) != 0) {
                //Count rows to show user after upload
                $totalUploadedCnt = 0;
                $totalSkippedCnt = 0;
                //Get new customer to save
                $customerList = [];

                //Read two sheets
                foreach ($uploadFileDataList as $index => $sheets) {
                    //header
                    $first = true;

                    foreach ($sheets as $item => $data) {
                        //check if header has more than 21 column
                        if ($first) {
                            // if 22 column has a value, let user knows that extra column exists
                            if (!empty($data[22])) {
                                throw new \Exception("MoreColumn");
                            } else {
                                $first = false;
                                continue;
                            }
                        }

                        //Get data under header
                        //Fields Validation
                        if (empty($data[0]) && empty($data[1]) && empty($data[2]) && empty($data[3]) && empty($data[4]) && empty($data[5])) {
                            continue;
                        } else {
                            if (empty($data[0]) || empty($data[1]) || empty($data[2]) || empty($data[3]) || empty($data[4]) || empty($data[5])) {
                                throw new \Exception("Missing");
                            } else {
                                //add customer info to array to save new customers
                                array_push($customerList, [
                                    'customer_code' => trim($data[1]),
                                    'customer_name' => trim($data[2]),
                                ]);

                                $sourceHeaderNo = trim($data[4]);
                                $serialNo = trim($data[12]);
                                $sourceLineNo = (int)$data[5];

                                //check duplicated item with same rma order#, line#, serial#
                                $checkRmaOrderEvent = $connection
                                    ->table(DB::raw('receive_details as rd'))
                                    ->select(DB::raw(' /* SQL_ID: ReportServiceImpl-uploadExcelForRmaOrderItemBulkInsert-checkRMAOrderEvent */
                                        rc.source_header_no,
                                        rd.source_line_no
                                    '))
                                    ->join('receive_contents AS rc', function ($join) {
                                        $join->on('rc.receive_id', 'rd.receive_id')
                                            ->whereRaw('klrc.del_yn = \'N\'');
                                    })
                                    ->where('rd.del_yn', 'N')
                                    ->where('rc.source_header_no', $sourceHeaderNo)
                                    ->where('rd.source_line_no', $sourceLineNo)
                                    ->where('rd.serial_no', $serialNo)
                                    ->get();

                                //number of RMA order list
                                $modelCount = $checkRmaOrderEvent->count();

                                if ($modelCount == 0) {
                                    DB::beginTransaction();
                                    //insert
                                    $currReceiveId = $connection
                                        ->table('receive_contents')
                                        ->insertGetId([
                                            'transaction_date' => trim($data[0]),
                                            'customer_code' => trim($data[1]),
                                            'customer_name' => trim($data[2]),
                                            'return_type' => trim($data[3]),
                                            'source_header_id' => trim($data[6]),
                                            'source_header_no' => $sourceHeaderNo,
                                            'business_id' => $sourceHeaderNo,
                                            'ship_to_city' => trim($data[20]),
                                            'ship_to_state' => trim($data[21]),
                                            'legal_entity_name' => 'US',
                                            'cancel_flag' => 'N',
                                            'edi_master_id' => 0,
                                            'source_type_code' => 9,
                                            'status' => $newStatusRmaOrder,
                                            'created_user' => $userCode
                                        ]);

                                    $connection
                                        ->table('receive_details')
                                        ->insertGetId([
                                            'receive_id' => $currReceiveId,
                                            'pick_seq_no' => 1,
                                            'source_line_id' => 1,
                                            'source_line_no' => $sourceLineNo,
                                            'item_code' => trim($data[10]),
                                            'order_qty' => trim($data[11]),
                                            'serial_no' => $serialNo,
                                            'organization_code' => trim($data[13]),
                                            'subinventory_code' => trim($data[14]),
                                            'cancel_flag' => 'N',
                                            'status' => $newStatusRmaOrder,
                                            'created_user' => $userCode
                                        ]);
                                    $totalUploadedCnt++;
                                    DB::commit();

                                } else {
                                    $totalSkippedCnt++;
                                    continue;
                                }
                            }
                        }
                    }
                }

                $customers = array();
                //remove duplicate $customerList
                foreach ($customerList as $code => $value) {
                    $customers[$value['customer_code']] = $value['customer_name'];
                }

                //if customer is new, save it
                foreach ($customers as $customerCode => $customerName) {
                    //check customercode exists
                    $checkCustomerCodeEvent = $connection
                        ->table(DB::raw('customer'))
                        ->select(DB::raw('customer_code, customer_name'))
                        ->where('del_yn', 'N')
                        ->where('customer_code', $customerCode)
                        ->get();

                    //number of RMA order list
                    $customerCount = $checkCustomerCodeEvent->count();

                    if ($customerCount == 0) {
                        $connection
                            ->table('customer')
                            ->insertGetId([
                                'customer_code' => $customerCode,
                                'customer_name' => $customerName,
                                'created_user' => 'bulk',
                                'updated_user' => 'bulk'
                            ]);
                    }
                }
            } else {
                throw new \Exception("NoData");
            }

            return array('status_code' => 200, 'success_data'=> 'Total Uploaded Count:' .$totalUploadedCnt .', Total Skipped Count:'. $totalSkippedCnt);

        } catch (\Exception $exception) {
            DB::rollBack();
            if ($exception->getMessage() == "Missing") {
                return array('error' => ($index + 1) . ' Sheet: Row ' . ($item + 5) . ' is missing required fields.');
            } else if ($exception->getMessage() == "NoData") {
                return array('error' => 'Excel Template Data Empty.');
            } else if ($exception->getMessage() == "MoreColumn") {
                return array('error' => ($index + 1) . ' Sheet: has extra column(s). Please check again.');
            }
            return array('error' => 'Please check template excel file.' . $exception);
        }
    }
}
