'use strict';

(function () {

    var email;
    var password;
    var client_id;
    var client_url;
    var _settings;
    var currentMailItem;
    var re_email = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/ ;
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
        _settings = Office.context.roamingSettings;
        currentMailItem = Office.context.mailbox.item;  
        $('.content-sender-display-name').text(Office.context.mailbox.item.from.displayName);
        
        // logout

        // Check if settings exist
        if(_settings.get("user_email")){
            // Get user details
            email = _settings.get("user_email");
            password = _settings.get("user_password");
            client_id = _settings.get("user_client_id");
            client_url = 'https://' + client_id + ".salesseek.net/api";
            login(client_id, email, password, client_url);
        }else{
            // provide authentication window
            $('.page-div').addClass('hidden')
            $('#div-authentication-details').removeClass('hidden');
        }
        
    });
  };

  $('#btn-login').click(function(event){
    
      if(!validateLoginForm())
        return;

      // Login form validated
      // proceed to retrieve values and log in
      client_id = $('#field-account-name').val();
      email = $('#field-email').val();
      password = $('#field-password').val();
      client_url = client_url = 'https://' + client_id + ".salesseek.net/api";
      login(client_id, email, password, client_url);
  })

  function logout(){
    _settings.remove("user_email");
    _settings.remove("user_password");
    _settings.remove("user_client_id");
    _settings.saveAsync(saveAppCredentialsSettingsCallback);
  }

  function login(client_id, p_email, p_password, p_client_url){
      // Try authentication
      $.ajax(
        {
        type: 'POST',
        url: p_client_url + '/login',
        crossDomain: true,
        xhrFields: {
            withCredentials: true
        },
        data: {
            email_address: p_email,
            password: p_password
        },
        dataType: 'json',
        success: function(data, textStatus, request) {
            // Store details to roaming settings
            _settings.set("user_email", p_email);
            _settings.set("user_password", p_password);
            _settings.set("user_client_id", client_id);
            _settings.saveAsync(saveAppCredentialsSettingsCallback);  
        },
        error: function() {
            // Remove submitting messaging
            showAlert('error', "Login error. Check your credentials and try agaon.", 5000, '#div-authentication-details');
            $('#log-message').text(error.responseJSON.detail + ", StatusCode: " + error.status + ", headers: " + error.getAllResponseHeaders);
        }
    });

  }

  function saveAppCredentialsSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
        
    }else{
        searchByEmail(Office.context.mailbox.item.from.emailAddress)
    }
}

  function validateLoginForm(){
      var validated = true;

      // Check required
      $.each($('#div-authentication-details input'), function(index, inputObject){
          if($.trim($(inputObject).val()) == ""){
              // Missing value... Note
              $(inputObject).siblings('.error-message').html($(inputObject).data('name') + " is required.");
              validated = false;
          }
          else{
            $(inputObject).siblings('.error-message').val("");
          }
      })
      // check email format
      
      var tempEmail = $.trim($('#field-email').val());
      if(tempEmail != "" && !tempEmail.match(re_email)){
        $('#field-email').siblings('.error-message').val("Enter valid email address.");
        validated = false;
      }
      return validated;
  }

  function searchByEmail(email_address){

    // Close all other views
    $('.page-div').addClass('hidden');
    $.ajax(
        {
            type: 'GET',
            url: client_url + '/search?terms=' + email_address,
            crossDomain: true,
            xhrFields: {
                withCredentials: true
            },
            data: {
                email_address: email,
                password: password
            },
            dataType: 'json',
            success: function(data, textStatus, request) {

                // Check if result has type individual
                
                if(data.batch_size > 0){
                    // Individual found flag
                    var individual_found = false;
                    // Some data retrieved... Process this data
                    $('#log-message').text("Batch size: " + data.batch_size + ", typeof: " + typeof(data.results));
                    $.each(data.results, function(index, result){
                        $('#log-message').text("Iterations: Result.type: " + result.type);
                        if(result.type == "salesseek.contacts.models.individual.Individual"){ 
                            // Individual found... We only process the first matching for the moment
                            individual_found = true;
                            // Process individual details
                            $('#content-organization-name').text(result.item.organization.name);
                            $('#content-organization-role').text(result.item.role);
                            // Communication details
                            var org_websites = "";
                            var org_emails = "";
                            var org_phones = "";
                            var org_linkedin = "";
                            var org_twitter = "";
                            var org_facebook = "";
                            // var org_googleplus = "";
                            $.each(result.item.organization.communication, function(commIndex, comm){
                                if(comm.name == "website"){
                                    if(org_websites != "")
                                        org_websites += ", ";
                                    org_websites += comm.value;
                                }
                                else if(comm.name == "phone"){
                                    if(org_phones != "")
                                    org_phones += ", ";
                                    org_phones += comm.value;
                                }
                                else if(comm.name == "email"){
                                    if(org_emails != "")
                                    org_emails += ", ";
                                    org_emails += comm.value;
                                }
                                else if(comm.name == "linkedin"){
                                    if(org_linkedin != "")
                                    org_linkedin += ", ";
                                    org_linkedin += comm.value;
                                }
                                else if(comm.name == "twitter"){
                                    if(org_twitter != "")
                                    org_twitter += ", ";
                                    org_twitter += comm.value;
                                }
                                else if(comm.name == "facebook"){
                                    if(org_facebook != "")
                                    org_facebook += ", ";
                                    org_facebook += comm.value;
                                }
                                // else if(comm.name == "googleplus"){
                                //     if(org_googleplus != "")
                                //     org_googleplus += ", ";
                                //     org_googleplus += comm.value;
                                // }
                            });

                            // Add values Responsibly...
                            $('#content-organization-website').text(org_websites);

                            // Adding email
                            $('._value_v0gc7_56.email').text(org_emails);
                            $('._value_v0gc7_56.email').attr('href', "mailto:" + org_emails);
                            if(org_emails == "")
                                $('._value_v0gc7_56.email').closest('._FieldContainer_13z8k_1').addClass('hidden');
                            else
                                $('._value_v0gc7_56.email').closest('._FieldContainer_13z8k_1').removeClass('hidden');

                            // Adding phone
                            $('._value_v0gc7_56.phone').text(org_phones);
                            $('._value_v0gc7_56.phone').attr('href', "tel:" + org_phones);
                            if(org_phones == "")
                                $('._value_v0gc7_56.phone').closest('._FieldContainer_13z8k_1').addClass('hidden');
                            else
                                $('._value_v0gc7_56.phone').closest('._FieldContainer_13z8k_1').removeClass('hidden');

                            // Adding organization name
                            $('#log-message').text("Organization: " + result.item.organization.name);

                            // Adding Twiter
                            $('._value_v0gc7_56.twitter').text("@" + org_twitter);
                            $('._value_v0gc7_56.twitter').attr('href', "https://twitter.com/" + org_twitter);
                            if(org_twitter == "")
                                $('._value_v0gc7_56.twitter').closest('._FieldContainer_13z8k_1').addClass('hidden');
                            else
                                $('._value_v0gc7_56.twitter').closest('._FieldContainer_13z8k_1').removeClass('hidden');

                            // Adding Facebook
                            $('._value_v0gc7_56.facebook').text(org_facebook);
                            $('._value_v0gc7_56.facebook').attr('href', org_facebook);
                            if(org_facebook == "")
                                $('._value_v0gc7_56.facebook').closest('._FieldContainer_13z8k_1').addClass('hidden');
                            else
                                $('._value_v0gc7_56.facebook').closest('._FieldContainer_13z8k_1').removeClass('hidden');

                            // Adding LinkedIn
                            $('._value_v0gc7_56.linkedin').text(org_linkedin);
                            $('._value_v0gc7_56.linkedin').attr('href', org_linkedin);
                            if(org_linkedin == "")
                                $('._value_v0gc7_56.linkedin').closest('._FieldContainer_13z8k_1').addClass('hidden');
                            else
                                $('._value_v0gc7_56.linkedin').closest('._FieldContainer_13z8k_1').removeClass('hidden');

                            // Hide addition section
                            // Show organization section
                            $('#div-organization-details').removeClass('hidden');
                            $('#div-add-user-to-crm').addClass('hidden');
                        }
                    })
                    if(!individual_found){
                        // individual not found
                        promptUserAddition();
                    }
                }else{
                    // No data retrieved... Start process to
                    // 1. Inform user
                    // 2. Allow for entry of new user into SalesSeek
                    promptUserAddition();
                }
                //$('#log-message').text("Data for: " + email_address + "-------------------------" + JSON.stringify(data));
            },
            error: function() {
                // Remove submitting messaging
                $('#log-message').text("Error");
                $('#log-message').text(error.responseJSON.detail + ", StatusCode: " + error.status + ", headers: " + error.getAllResponseHeaders);
            }
        }
    );
  }
    
  function promptUserAddition(){
    $('#div-organization-details').addClass('hidden');
    $('#div-add-user-to-crm').removeClass('hidden');
    
    $.ajax(
        {
        type: 'GET',
        url: client_url + '/organizations',
        crossDomain: true,
        xhrFields: {
            withCredentials: true
        },
        data: {
            email_address: email,
            password: password
        },
        dataType: 'json',
        success: function(data, textStatus, request) {
            $('#log-message').text("Final success");
            var org_select = $('#field-select-organization');
            $.each(data, function(index, organization){
                // Add option to select
                org_select.append('<option value="' + organization.id + '">' + organization.name + '</option>');
            })
        },
        error: function() {
            // Remove submitting messaging
            $('#log-message').text("Error");
            $('#log-message').text(error.responseJSON.detail + ", StatusCode: " + error.status + ", headers: " + error.getAllResponseHeaders);
        }
    });

    // Auto populate names
    if(currentMailItem != null){
        var nameParts = currentMailItem.from.displayName.split(" ");
        if(nameParts.length > 0){
            $('.field-first-name').val(nameParts[0]);
        }

        if(nameParts.length > 1){
            $('.field-last-name').val(nameParts[1]);
        }
    }
  }

  $('#field-check-organization').change(function(event){
    $('.field-input-organization').val("");
    if( $(this).is(':checked') ) {
        $('.field-input-organization').attr('disabled',false).focus();
        $('#field-select-organization').val("").attr('disabled', 'disabled');
        $('#field-select-organization').siblings('.error-message').html("");
        
    }else{
        $('.field-input-organization').attr('disabled', 'disabled');
        $('#field-input-organization').siblings('.error-message').html("");
        $('#field-select-organization').val("").attr('disabled', false).focus();        
    }
    
  })

  $('.checkbox').click(function(event){
        $('#field-check-organization').prop('checked', $('#field-check-organization').is(':checked') ? false: true);
        $('#field-check-organization').trigger('change');
  })

  $('.field-name').change(function(event){
        // Get all with name and check if empty
        validateNames();
  })

  $('.field-auth').change(function(event){
    if($.trim($(this).val()) == ""){
        // Missing value... Note
        $(this).siblings('.error-message').html($(this).data('name') + " is required.");
        validated = false;
    }
    else{
      $(this).siblings('.error-message').val("");
    }

    if($(this).attr('id') == "field-email"){
      var tempEmail = $.trim($(this).val());
      if(tempEmail != "" && !tempEmail.match(re_email)){
        $(this).siblings('.error-message').html("Enter valid email address.");
        validated = false;
      }
    }
  })

  $('#field-account-name').keyup(function(event){
      // Update content on key press
      $('#help-field-account-name').text($(this).val());
  })

  $('.field-organization').change(function(event){
        validateOrganization();
  })

  function validateNames(){
        var missingCount = 0;
        $.each($('.field-name'), function(index, nameInput){
            if($.trim($(nameInput).val()) == ""){
                $(this).siblings('.error-message').html($(this).data('name') + ' is required.')
                missingCount++;
            }else{
                $(this).siblings('.error-message').html("");
            }
        })

        return missingCount == 0;
  }

  function validateOrganization(){
        // Check if new is checked
        var validated = true;
        
        if($('#field-check-organization').is(':checked')){
            // Validate new organization entry
            if($.trim($('#field-input-organization').val()) == ""){
                validated = false
                $('#field-input-organization').siblings('.error-message').html("New organization name is required.");
            }else{
                $('#field-input-organization').siblings('.error-message').html("");
            }

            $('#field-select-organization').siblings('.error-message').html("");
        }else{
            // Validate organization selection
            if($.trim($('#field-select-organization').val()) == ""){
                validated = false;
                $('#field-select-organization').siblings('.error-message').html("Organization is required.");
            }else{
                $('#field-select-organization').siblings('.error-message').html("");
            }
        }

        return validated;
  }

  function validateAddIndividualForm(){
        if(validateNames() && validateOrganization()){
            return true;
        }else{
            return false;
        }
  }

  $('#btn-save-individual').click(function(event){

        // Trying to save a user
        if(!validateAddIndividualForm())
            return;

        // Everything validates...
        // Check if new organization...
        $('#loading-gif-add-individual').removeClass('hidden');
        $('#loading-gif-alt').addClass('hidden');
        
        $('#div-add-user-to-crm .detail-section').attr('disabled', 'disabled');
        if($('#field-check-organization').is(':checked')){
            // Create a new organization then add user
            var data = JSON.stringify({"name": $('#field-input-organization').val()});
            $.ajax(
                {
                type: 'POST',
                url: client_url + '/organizations',
                crossDomain: true,
                xhrFields: {
                    withCredentials: true
                },
                data: data,
                dataType: 'json',
                success: function(data, textStatus, request) {
                    // Returns organization.. Need to create user next..
                    createUserAndAddToOrganization(data.id);
                },
                error: function() {
                    // Remove submitting messaging
                    $('#loading-gif-add-individual').addClass('hidden');
                    $('#loading-gif-alt').removeClass('hidden');
                    $('#div-add-user-to-crm .detail-section').attr('disabled', false);

                    showAlert('error', "Could not add " + currentMailItem.from.displayName + " to SalesSleek. Please try again or contact system administrator.", 5000, '#div-add-user-to-crm');
                }
            });
        }else{
            // Add user to existing organization
            createUserAndAddToOrganization($('#field-select-organization').val());
        }
  });

  function createUserAndAddToOrganization(organization_id){
        // Get user details..
        var firstName = $('#field-first-name').val();
        var lastName = $('#field-last-name').val();
        var email_address = currentMailItem.from.emailAddress;
        var role = $('#field-role').val();
        var data = JSON.stringify({
            "first_name": firstName,
            "last_name": lastName, 
            "communication": [
                {
                    "name": "Work",
                    "medium": "email",
                    "value": email_address,
                    "comments": ""
                }
            ],
            "organization_id": organization_id,
            "roles": [
                {
                    "title": role,
                    "organization_id": organization_id
                }
            ]   
        });

        $.ajax(
            {
            type: 'POST',
            url: client_url + '/individuals',
            crossDomain: true,
            xhrFields: {
                withCredentials: true
            },
            data: data,
            dataType: 'json',
            success: function(data, textStatus, request) {
                // Returns organization.. Need to create user next..
                $('#loading-gif-add-individual').addClass('hidden');
                $('#loading-gif-alt').removeClass('hidden');
                $('#div-add-user-to-crm .detail-section').attr('disabled', false);

                $('#div-add-user-to-crm  .alert-message').text(currentMailItem.from.displayName + " successfully added to SalesSleek.");

                showAlert('success', currentMailItem.from.displayName + " successfully added to SalesSleek.", 2000, '#div-add-user-to-crm', searchByEmail(currentMailItem.from.emailAddress));
            },
            error: function() {
                // Remove submitting messaging
                $('#loading-gif-add-individual').addClass('hidden');
                $('#loading-gif-alt').removeClass('hidden');
                $('#div-add-user-to-crm .detail-section').attr('disabled', false);

                showAlert('error', "Could not add " + currentMailItem.from.displayName + " to SalesSleek. Please try again or contact system administrator.", 5000, '#div-add-user-to-crm');
            }
        });

    }

    function showAlert(type, message, displayTime, alertSectionId, callback){
        if(type == "error")
            $(alertSectionId + ' .alert-message').css('color', '#dc3545');
        else if(type == "success")
            $(alertSectionId + ' .alert-message').css('color', '#28a745');

        $(alertSectionId + ' .alert-message').text( message );    
        setTimeout(function(){
            $(alertSectionId + ' .alert-message').text(""); // Clear text after 5 seconds
            if(callback != null){
                callback();
            }
        }, displayTime);
    }


})();