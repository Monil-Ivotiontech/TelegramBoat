export enum ApiMethod {
    GET = "GET",
    POST = "POST",
    PUT = "PUT",
    DELETE = "DELETE"
}

export enum EndPoints {
    // Auth API's
    // register
    register_admin_user = "be_api/admin/auth/register_admin_user",

    // login 
    login_admin_request = "be_api/admin/auth/login_admin_request", //done
    login_admin_success = "be_api/admin/auth/login_admin_success", //done
    login_admin_cookies = "be_api/admin/auth/login_admin_cookies", //done
    cookies_json = "json",
    get_bucket_folder_name = "be_api/admin/auth/get_bucket_folder_name",

    // logout
    admin_logout = "be_api/admin/auth/admin_logout", //done

    // credentials
    admin_change_password = "be_api/admin/auth/admin_change_password",
    admin_forgot_password = "be_api/admin/auth/admin_forgot_password",
    admin_reset_password = "be_api/admin/auth/admin_reset_password",

    // Admin API's
    get_registered_admin_list = "be_api/admin/auth/get_registered_admin_list",
    delete_registered_admin_user = "be_api/admin/auth/delete_registered_admin_user",

    // MT5 API's
    get_mt5_manager_detail = "be_api/admin/auth/get_mt5_manager_detail",
    get_mt5_group_list_for_selected_manager = "be_api/admin/auth/get_mt5_group_list_for_selected_manager",
    get_mt5_user_list_for_selected_manager = "be_api/admin/auth/get_mt5_user_list_for_selected_manager",
    clone_mt5_manager_to_new = "be_api/admin/auth/clone_mt5_manager_to_new",
    save_mt5_manager_chatid = "be_api/admin/auth/save_mt5_manager_chatid",
    get_mt5_manager_chatid = "be_api/admin/auth/get_mt5_manager_chatid",
    update_mt5_manager_chatid = "be_api/admin/auth/update_mt5_manager_chatid",

    get_mt5_group_detail = "be_api/admin/auth/get_mt5_group_detail",
    get_mt5_groupwise_user_count = "be_api/admin/auth/get_mt5_groupwise_user_count",
    get_mt5_groupwise_user_detail = "be_api/admin/auth/get_mt5_groupwise_user_detail",
    clone_mt5_group_to_new_group = "be_api/admin/auth/clone_mt5_group_to_new_group",

    get_mt5_user_detail = "be_api/admin/auth/get_mt5_user_detail",
    create_mt5_dummy_user = "be_api/admin/auth/create_mt5_dummy_user",
    create_mt5_dummy_client = "be_api/admin/auth/create_mt5_dummy_client", //pending
    create_mt5_new_or_clone_user = "be_api/admin/auth/create_mt5_new_or_clone_user",
    update_mt5_user = "be_api/admin/auth/update_mt5_user",

    save_mt5_manager_bot_access_detail = "be_api/admin/auth/save_mt5_manager_bot_access_detail",
    get_mt5_manager_bot_access_detail = "be_api/admin/auth/get_mt5_manager_bot_access_detail",
    update_mt5_manager_bot_access_detail = "be_api/admin/auth/update_mt5_manager_bot_access_detail",
    delete_mt5_manager_bot_access_detail = "be_api/admin/auth/delete_mt5_manager_bot_access_detail",

    save_venus_manager_bot_access_detail = "be_api/admin/auth/save_venus_manager_bot_access_detail",
    get_venus_manager_bot_access_detail = "be_api/admin/auth/get_venus_manager_bot_access_detail",
    update_venus_manager_bot_access_detail = "be_api/admin/auth/update_venus_manager_bot_access_detail",
    delete_venus_manager_bot_access_detail = "be_api/admin/auth/delete_venus_manager_bot_access_detail",

    // Venus Group API's
    save_venus_group_detail = "be_api/admin/auth/save_venus_group_detail",
    get_venus_group_detail = "be_api/admin/auth/get_venus_group_detail",
    update_venus_group_detail = "be_api/admin/auth/update_venus_group_detail",
    delete_venus_group_detail = "be_api/admin/auth/delete_venus_group_detail",
    save_venus_group_mapping_detail = "be_api/admin/auth/save_venus_group_mapping_detail",
    get_venus_group_mapping_detail = "be_api/admin/auth/get_venus_group_mapping_detail",
    delete_user_from_venus_group_mapping = "be_api/admin/auth/delete_user_from_venus_group_mapping",

    // Venus Manager API's
    save_venus_manager_detail = "be_api/admin/auth/save_venus_manager_detail",
    get_venus_manager_detail = "be_api/admin/auth/get_venus_manager_detail",
    update_venus_manager_detail = "be_api/admin/auth/update_venus_manager_detail",
    delete_venus_manager_detail = "be_api/admin/auth/delete_venus_manager_detail",
    get_venus_manager_mapping_detail = "be_api/admin/auth/get_venus_manager_mapping_detail",
    save_venus_manager_mapping_detail = "be_api/admin/auth/save_venus_manager_mapping_detail",
    delete_user_from_venus_manager_mapping = "be_api/admin/auth/delete_user_from_venus_manager_mapping",
    save_venus_manager_group_mapping_detail = "be_api/admin/auth/save_venus_manager_group_mapping_detail",
    get_venus_manager_group_mapping_detail = "be_api/admin/auth/get_venus_manager_group_mapping_detail",
    delete_group_from_venus_manager_group_mapping = "be_api/admin/auth/delete_group_from_venus_manager_group_mapping",


    // Command Name API's   
    save_venus_command_detail = "be_api/admin/auth/save_venus_command_detail",
    get_venus_command_detail = "be_api/admin/auth/get_venus_command_detail",
    delete_venus_command_detail = "be_api/admin/auth/delete_venus_command_detail",
    update_venus_command_detail = "be_api/admin/auth/update_venus_command_detail",

    // Command Mapping API's 
    get_venus_command_list_for_mapping = "be_api/admin/auth/get_venus_command_list_for_mapping",
    get_mt5_command_list_for_mapping = "be_api/admin/auth/get_mt5_command_list_for_mapping",
    save_venus_command_mapping_detail = "be_api/admin/auth/save_venus_command_mapping_detail",
    get_venus_command_mapping_detail = "be_api/admin/auth/get_venus_command_mapping_detail",
    delete_group_from_venus_command_mapping = "be_api/admin/auth/delete_group_from_venus_command_mapping",
    delete_manager_from_venus_command_mapping = "be_api/admin/auth/delete_manager_from_venus_command_mapping",
    get_result_for_m2m_command = "be_api/admin/auth/get_result_for_m2m_command",
    get_result_for_com_pos_command = "be_api/admin/auth/get_result_for_com_pos_command",
    get_result_for_total_pos_command = "be_api/admin/auth/get_result_for_total_pos_command",
    get_result_for_update_all_command = "be_api/admin/auth/get_result_for_update_all_command",

    // User Trading API's
    get_mt5_groupwise_user_deal_detail = "be_api/admin/auth/get_mt5_groupwise_user_deal_detail",
    get_mt5_groupwise_userwise_deal_summary = "be_api/admin/auth/get_mt5_groupwise_userwise_deal_summary",
    get_mt5_groupwise_user_position_detail = "be_api/admin/auth/get_mt5_groupwise_user_position_detail",

    // Duplicate IP API's
    get_data_duplication_analysis = "be_api/admin/auth/get_data_duplication_analysis",
    generate_data_summary_analysis = "be_api/admin/auth/generate_data_summary_analysis",

    // Symbol API's
    save_symbol_category_detail = "be_api/admin/auth/save_symbol_category_detail",
    get_symbol_category_detail = "be_api/admin/auth/get_symbol_category_detail",
    update_symbol_category_detail = "be_api/admin/auth/update_symbol_category_detail",
    delete_symbol_category_detail = "be_api/admin/auth/delete_symbol_category_detail",
}

export enum UtilVariable {
    ENCRYPT_KEY = "ArtTTTd22457fjklllMKKDERYU2358899MDaTyupe987R",
    GEOAPI_KEY = "43a2a8591aa04a2181d83df24f76b2d9",
}