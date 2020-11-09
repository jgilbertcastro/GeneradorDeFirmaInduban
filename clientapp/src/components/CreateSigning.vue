<template>
    <div class="CreateSigning cspio">
        <div id="cspio-page">
            <div id="cspio-content">

                <img id="cspio-logo" src="../assets/logo.png">

                <h1 id="cspio-headline">Generador de Firma</h1>

                <div id="cspio-form">
                    <div id="cspio-field-wrapper">
                        <div class="row">
                            <div class="col-md-12 seperate">
                                <div class="input-group">
                                    <input id="cspio-email" name="employeeCode" class="form-control input-lg form-el" type="text"  v-model="employeeCode" placeholder="Codigo de empelado" required />
                                    <span class="input-group-btn"><button id="cspio-subscribe-btn" type="button" class="btn btn-lg btn-primary form-el noglow" @click="generate">Generar</button></span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <div id="cspio-socialprofiles">
                    <a href="https://www.facebook.com/IndubanSAS" target="_blank"><i class="fa fa-facebook-official fa-2x"></i></a>
                    <a href="https://www.instagram.com/indubansas/" target="_blank"><i class="fa fa-instagram fa-2x"></i></a>
                    <a href="https://twitter.com/IndubanRD" target="_blank"><i class="fa fa-twitter fa-2x"></i></a>
                </div>

            </div><!-- end of #cspio-content -->
        </div>
    </div>
</template>

<script>
    import Axios from "axios";
    import Swal from 'sweetalert2';

    export default {
        name: 'CreateSigning',
        data() {
            return {
                employeeCode: ""
            }
        },
        methods: {
            generate() {
                var self = this;
                Axios.get('GenerateSigning/generate?employeeCode=' + self.employeeCode).then(function (data) {
                    if (data.data != null) {
                        window.location.href = "GenerateSigning/downloadfile?employeeCode=" + data.data.id;
                    } else {
                        Swal.fire({
                            icon: 'error',
                            title: 'Oops...',
                            text: 'Verifique su codigo y trate nuevamente.'
                        });
                    }
                }).catch(function (error) {
                    Swal.fire({
                        icon: 'error',
                        title: 'Ha ocurrido un error, contacte el departamento de tecnologia',
                        text: error
                    });
                });
            }
        }
    }
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
    .CreateSigning {
        height: 100%;
        /*background: #ffffff url(../assets/background.png);
        no-repeat: center bottom fixed;
        -webkit-background-size: cover;
        -moz-background-size: cover;
        -o-background-size: cover;
        background-size: cover;*/
    }

    #cspio-page {
        background-color: rgba(0,0,0,0);
    }

    .flexbox #cspio-page {
        align-items: center;
        justify-content: center;
    }

    .cspio body {
        background: transparent;
    }

        .cspio body, .cspio body p {
            font-family: Helvetica, Arial, sans-serif;
            font-weight: 400;
            font-style:;
            font-size: 20px;
            line-height: 1.50em;
            color: #ffffff;
        }

    ::-webkit-input-placeholder {
        font-family: Helvetica, Arial, sans-serif;
        font-weight: 400;
    }

    ::-moz-placeholder {
        font-family: Helvetica, Arial, sans-serif;
        font-weight: 400;
    }

    :-ms-input-placeholder {
        font-family: Helvetica, Arial, sans-serif;
        font-weight: 400;
    }

    :-moz-placeholder {
        font-family: Helvetica, Arial, sans-serif;
        font-weight: 400;
        font-style:;
    }

    .cspio h1, .cspio h2, .cspio h3, .cspio h4, .cspio h5, .cspio h6 {
        font-family: 'Pacifico';
        color: #ffffff;
    }

    #cspio-headline {
        font-family: 'Pacifico';
        font-weight: 400;
        font-size: 48px;
        color: #ffffff;
        line-height: 1.00em;
    }

    .cspio button {
        font-family: Helvetica, Arial, sans-serif;
        font-weight: 400;
    }

    .cspio a, .cspio a:visited, .cspio a:hover, .cspio a:active {
        color: #ffffff;
    }

    #cspio-socialprofiles a {
        color: #ffffff;
    }

    .cspio .btn-primary,
    .cspio .btn-primary:focus,
    .gform_button,
    #mc-embedded-subscribe, .submit-button {
        color: black;
        text-shadow: 0 -1px 0 rgba(255,255,255,0.3);
        background-color: #ffffff;
        background-image: -moz-linear-gradient(top,#ffffff,#d9d9d9);
        background-image: -ms-linear-gradient(top,#ffffff,#d9d9d9);
        background-image: -webkit-gradient(linear,0 0,0 100%,from(#ffffff),to(#d9d9d9));
        background-image: -webkit-linear-gradient(top,#ffffff,#d9d9d9);
        background-image: -o-linear-gradient(top,#ffffff,#d9d9d9);
        background-image: linear-gradient(top,#ffffff,#d9d9d9);
        background-repeat: repeat-x;
        filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#ffffff', endColorstr='#d9d9d9', GradientType=0);
        border-color: #d9d9d9 #d9d9d9 #b3b3b3;
        border-color: rgba(0,0,0,0.1) rgba(0,0,0,0.1) rgba(0,0,0,0.25);
        *background-color: #d9d9d9;
        filter: progid:DXImageTransform.Microsoft.gradient(enabled = false);
    }

        .cspio .btn-primary:hover,
        .cspio .btn-primary:active,
        .cspio .btn-primary.active,
        .cspio .btn-primary.disabled,
        .cspio .btn-primary[disabled],
        .cspio .btn-primary:focus:hover,
        .cspio .btn-primary:focus:active,
        .cspio .btn-primary:focus.active,
        .cspio .btn-primary:focus.disabled,
        .cspio .btn-primary:focus[disabled],
        #mc-embedded-subscribe:hover,
        #mc-embedded-subscribe:active,
        #mc-embedded-subscribe.active,
        #mc-embedded-subscribe.disabled,
        #mc-embedded-subscribe[disabled] {
            background-color: #d9d9d9;
            *background-color: #cccccc;
        }

        .cspio .btn-primary:active,
        .cspio .btn-primary.active,
        .cspio .btn-primary:focus:active,
        .cspio .btn-primary:focus.active,
        .gform_button:active,
        .gform_button.active,
        #mc-embedded-subscribe:active,
        #mc-embedded-subscribe.active {
            background-color: #bfbfbf;
        }

    .form-control,
    .progress {
        background-color: rgba(255, 255, 255, 0.85);
    }

    #cspio-progressbar span,
    .countdown_section {
        color: black;
        text-shadow: 0 -1px 0 rgba(255,255,255,0.3);
    }

    .cspio .btn-primary:hover,
    .cspio .btn-primary:active {
        color: black;
        text-shadow: 0 -1px 0 rgba(255,255,255,0.3);
        border-color: #e6e6e6;
    }

    .cspio input[type='text']:focus {
        webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,0.075), 0 0 8px rgba(217,217,217,0.6);
        -moz-box-shadow: inset 0 1px 1px rgba(0,0,0,0.075), 0 0 8px rgba(217,217,217,0.6);
        box-shadow: inset 0 1px 1px rgba(0,0,0,0.075), 0 0 8px rgba(217,217,217,0.6);
    }

    #cspio-content {
        display: block;
        max-width: 600px;
        background-color: #000000;
        -webkit-border-radius: 2px;
        border-radius: 2px;
        -moz-background-clip: padding;
        -webkit-background-clip: padding-box;
        background-clip: padding-box;
        background-color: transparent;
    }

    .cspio .progress-bar,
    .countdown_section,
    .cspio .btn-primary,
    .cspio .btn-primary:focus,
    .gform_button {
        background-image: none;
        text-shadow: none;
    }

    .cspio input,
    .cspio input:focus {
        -webkit-box-shadow: none !important;
        box-shadow: none !important;
    }

    #cspio-page {
        background: -moz-radial-gradient(ellipse at center, rgba(0, 0, 0, 0.3) 0%,rgba(0, 0, 0, 0.2) 37%,rgba(0,0,0,0) 68%,rgba(0,0,0,0) 100%);
        background: -webkit-radial-gradient(ellipse at center, rgba(0, 0, 0, 0.3) 0%,rgba(0, 0, 0, 0.2) 37%,rgba(0,0,0,0) 68%,rgba(0,0,0,0) 100%);
        background: radial-gradient(ellipse at center, rgba(0, 0, 0, 0.3) 0%,rgba(0, 0, 0, 0.2) 37%,rgba(0,0,0,0) 68%,rgba(0,0,0,0) 100%);
    }

    .cspio body {
        background: -moz-radial-gradient(center, ellipse cover, rgba(0,0,0,0) 7%, rgba(0,0,0,0) 80%, rgba(0,0,0,0.23) 100%);
        background: -webkit-radial-gradient(center, ellipse cover, rgba(0,0,0,0) 7%,rgba(0,0,0,0) 80%,rgba(0,0,0,0.23) 100%);
        background: radial-gradient(ellipse at center, rgba(0,0,0,0) 7%,rgba(0,0,0,0) 80%,rgba(0,0,0,0.23) 100%);
        filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#00000000', endColorstr='#3b000000',GradientType=1 );
    }

    #cspio-subscribe-btn {
        background: transparent;
        border: 1px solid #fff !important;
        color: #fff;
    }

        #cspio-subscribe-btn:hover {
            background: rgba(255,255,255,0.2);
            color: #fff;
        }

    #cspio-credit img {
        margin-left: auto;
        margin-right: auto;
        width: 125px;
        margin-top: -4px;
    }

    #cspio-credit {
        font-size: 11px;
    }
</style>
